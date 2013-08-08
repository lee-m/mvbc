#Copyright (C) 2013 by Lee Millward
#
#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#
#The above copyright notice and this permission notice shall be included in
#all copies or substantial portions of the Software.
#
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
#THE SOFTWARE.

import sys;
import os;
import subprocess;

from sys import stdout, stdin;
from os import getcwd;
from os.path import abspath, isfile, join, basename;

#Location to the compiler executable
COMPILER_EXE_LOCATION = ".\\mvbc\\bin\\mvbc.exe";

#Possible directives that may appear within a file
COMPILER_OPTS_DIRECTIVE = "'#compiler-options:";
COMPILER_MESSAGE_DIRECTIVE = "'#compiler-message:";
COMPILER_EXIT_CODE_DIRECTIVE = "'#exit-code:";

#Test run execution success/failure 
class test_result:

    def __init__(self, fileName, result=True, failureMessage=None):

        self.file_name = fileName;
        self.result = result;

        if not self.result:
            self.message = "%s - FAIL: %s" % (basename(fileName), failureMessage);
        else:
            self.message = "%s - PASS" % basename(fileName);

#represents a single regression test to execute 
class regression_test:

    def __init__(self, fileName):
        
        #Lists of compiler options and expected output and the desired exit code
        self.file_name = fileName;
        self.compiler_options = None;
        self.compiler_messages = [];
        self.expected_exit_code = 0;

        seenCompilerOptionsDirective = False;
        seenCompilerExitCodeDirective = False;

        #Open the file and parse any directives
        with open(fileName, "r") as testFile:

            while True:

                line = testFile.readline();

                if line.startswith(COMPILER_OPTS_DIRECTIVE):

                    #Reject multiple compiler options directives
                    if seenCompilerOptionsDirective:
                        raise Exception("Multiple #compiler-options directives found in file %s" % fileName);
                    else:
                        seenCompilerOptionsDirective = True;
                        self.compiler_options = line.replace(COMPILER_OPTS_DIRECTIVE, "").strip().split(" ") + [self.file_name];

                elif line.startswith(COMPILER_MESSAGE_DIRECTIVE):

                    self.compiler_messages.append(line.replace(COMPILER_MESSAGE_DIRECTIVE, "").strip());

                elif line.startswith(COMPILER_EXIT_CODE_DIRECTIVE):
                
                    #Reject multiple compiler exit code directives
                    if seenCompilerExitCodeDirective:
                        raise Exception("Multiple #exit-code directives found in file %s" % fileName);
                    else:
                        seenCompilerExitCodeDirective = True;
                        self.expected_exit_code = int(line.replace(COMPILER_EXIT_CODE_DIRECTIVE, "").strip());

                else:

                    #End of the directives section
                    break;

    #Executes the test and checks the expected exit-code/expected messages.
    def execute_test(self):

        #Names of the temp files used to hold the stdout/stderr output
        stdOutRedirectFileName = self.file_name.replace(".vb", ".stdout");
        stdErrRedirectFileName = self.file_name.replace(".vb", ".stderr");

        try:

            #Create file handles for holding the stdout/stderr output
            stdOutHandle = open(stdOutRedirectFileName, "w+");
            stdErrHandle = open(stdOutRedirectFileName, "w+");

            #Run the compiler 
            proc = subprocess.Popen([os.path.abspath(COMPILER_EXE_LOCATION)] + self.compiler_options, stdout = stdOutHandle, stderr = stdErrHandle);
            exitCode = proc.wait();

            stdOutHandle.seek(0, 0);
            stdErrHandle.seek(0, 0);

            #Parse the output into a list of lines of text to check against the expected output
            #(stdout, stderr) = proc.communicate();
            receivedMessages = set([o.replace("mvbc : ", "").strip() for o in stdOutHandle.read().split("\n") if len(o) > 0]);

            if os.path.exists(stdErrRedirectFileName):
                receivedMessages = receivedMessages.union([e.replace("mvbc : ", "").strip() for e in stdErrHandle.read().split("\n") if len(e) > 0]);

            stdOutHandle.close();
            stdErrHandle.close();

            #Check for ICE messages first, one of these may result in different exit codes or a different number of messages
            #to be output
            iceMessageFound = False;

            for message in receivedMessages:

                if "Internal compiler error" in message:
                    iceMessageFound = True;
                    break;

            if iceMessageFound:
                return test_result(self.file_name, False, "Internal compiler error");

            #Check the exit code matches what we expect
            if exitCode != self.expected_exit_code:
                return test_result(self.file_name, False, "Expected exit code %d, got %d" % (self.expected_exit_code, exitCode));

            #Check we got the expected number of messages
            if len(receivedMessages) != len(self.compiler_messages):
                return test_result(self.file_name, False, "Expected %d messages, got %d" % (len(self.compiler_messages), len(receivedMessages)));
        
            #Check that the messages are the same
            if not receivedMessages.issubset(set(self.compiler_messages)):
                return test_result(self.file_name, False, "Message contents differ");
         
            return test_result(self.file_name);

        finally:
            
            if os.path.exists(stdOutRedirectFileName):
                os.remove(stdOutRedirectFileName);

            if os.path.exists(stdErrRedirectFileName):
                os.remove(stdErrRedirectFileName);

#Script entry point
if __name__ == "__main__":

    #Location of the regression tests folder and compiler executable
    testsPath =  join(getcwd(), "tests");

    #Enumerate all of the tests within the tests folder and any sub-folders within it and 
    #create a regresstion test instance for them
    executedTests = 0;
    tests = []

    for root, subfolders, files in os.walk(testsPath):
        tests += [regression_test(join(root, f)) for f in files];

    numTests = len(tests);
    numTestsPassed = 0;

    stdout.write("Found %d tests\n" % numTests);

    #Execute each test
    for test in tests:

        executedTests += 1;
        stdout.writelines("Executing tests: %.2f%%   \r" % round((executedTests / float(numTests)) * 100, 2));
        stdout.flush();

        testResult = test.execute_test();
       
        #If the test fails, output details of the failure
        if not testResult.result:
            stdout.write("%s\n" % testResult.message);
        else:
            numTestsPassed += 1;

    #Special case no tests passing to avoid division by 0
    if numTestsPassed == 0:
        stdout.write("\n\nTest run finished - 0 of %d tests passed (0%%)\n" % numTests);
    else:
        stdout.write("\n\nTest run finished - %d of %d tests passed (%.2f%%)\n" % (numTestsPassed, numTests, round((float(numTestsPassed) / numTests) * 100, 2)));