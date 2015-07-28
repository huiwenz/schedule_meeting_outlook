/// <reference path="uia.d.ts" />
"use strict";

// This is the script for running the test, which does all the following steps:
// Note: Assume we run this script inside of each script folder

//////////////////////////////////////////////////////////////////// VARIABLES ////////////////////////////////////////////////////////////////////
var exec = require('child_process').exec,
    copy_manifest,
    del_logs,
    open_apps_outlook,
    open_apps_narrator,
    close_apps_outlook,
    close_apps_narrator,
    del_manifest;

var fs = require('fs')
	, file
	, byte_size = 256
	, readbytes = 0
	, lastLineFeed
	, lineArray
	, continueReading = true; // Make the decision to continue reading or not

//////////////////////////////////////////////////////////////////// HELPER FUNCTIONS ////////////////////////////////////////////////////////////////////
// Reference: http://stackoverflow.com/questions/11225001/reading-a-file-in-real-time-using-node-js

// cleanUp - The helper function for cleaning up after test ends.
function cleanUp() {

    // Exit Narrator and Outlook

    // The taskkill command sends the WM_CLOSE message and then it is up to the application
    // whether to properly close. 
    // /IM specifies the image name of the process to be terminated.

    // An 'exit' event will be emitted after the child process ends.
    // Source of that command:
    // http://superuser.com/questions/727724/properly-close-not-kill-programs-from-the-command-line

    //close_apps_narrator = exec('taskkill /IM narrator.exe >nul', // Don't output any errors from here
    //    function (error, stdout, stderr) {

    //        if (error !== null) {
    //            console.log('exec error: ' + error);
    //        }

    //    });

    close_apps_outlook = exec('taskkill /IM outlook.exe >nul', // Don't output any errors from here
    function (error, stdout, stderr) {

        if (error !== null) {
            console.log('exec error: ' + error);
        }

    });

    // Remove test.manifest
    del_manifest = exec('del test.manifest.json', // Don't output any errors from here
    function (error, stdout, stderr) {

        if (error !== null) {
            console.log('exec error: ' + error);
        }

    });


}

// readlogs - The helper function which can continuously read logs.txt file
function readlogs() {

    var stats = fs.fstatSync(file);

    if (stats.size < readbytes + 1) {
        console.log('Hehe I am much faster than your writer..! I will sleep for a while, I deserve it!');
        setTimeout(readlogs, 5000);
    } else {

        fs.read(file, new Buffer(byte_size), 0, byte_size, readbytes, processsome);

    }


}

// processsome - The helper function for readlogs which helps to parse the data that has been read
// into lines
function processsome(err, bytecount, buff) {

    lastLineFeed = buff.toString('utf-8', 0, bytecount).lastIndexOf('\n');

    if (lastLineFeed > -1) {

        // Split the buffer by line
        lineArray = buff.toString('utf-8', 0, bytecount).slice(0, lastLineFeed).split('\n');

        // Find each line and print them out
        for (var i = 0; i < lineArray.length; i++) {

            // Now we treat this as our report. Just output each line.
            // TODO: May want to check if each line is from test.js or app.js,
            // and only output the lines from test.js.
            console.log("This is a line: " + lineArray[i]);

            if (lineArray[i].indexOf("Done") > -1) { // When we read a "Done" in the log file, stop reading
                continueReading = false;
            }

        }

        // Set a new position to read from
        readbytes += lastLineFeed + 1;

    } else {
        // No complete lines were read
        readbytes += bytecount;
    }

    if (continueReading) {
        readlogs();

    } else {

        // Finding "Done" means the test has finished. Should do some cleaning up.
        cleanUp();

    }
}

// A helper function for changing focus to the Outlook window
function locateFocusOnApp() {

    return Q.fcall(function () {
        return uia.root().findFirst(function (el) { return (el.name.indexOf("Outlook") > -1); }, 0, 1);
    }).then(function (outlookWindow) {

        outlookWindow.setFocus();
        console.log("SET");

    }, function (error) {
        throw new Error("The promise for getting the Outlook window fails. ");
    })

}


// A helper function for checking the log file
function openLogFile() {

    // After both apps have opened, we start reading logs.txt CONTINUOUSLY
    // NOTE: May want to make the app to check the presence of logs.txt
    fs.open('../logs.txt', 'r', function (err, fd) {

        if (err !== null) {
            console.log('exec error: ' + err);
            console.log('Hmm no log files found!');
            setTimeout(openLogFile, 5000);
        } else {
            file = fd;
            readlogs();
        }


    });

}

//////////////////////////////////////////////////////////////////// STEPS ////////////////////////////////////////////////////////////////////
// Step 1: Copy test.manifest to make the test run

copy_manifest = exec('copy ..\\test.manifest.json',
    function (error, stdout, stderr) {

        if (error !== null) {
            console.log('exec error: ' + error);
        } else {

            // Step 2: Delete logs.txt, which is also located in NarratorScripts folder
            del_logs = exec('del ..\\logs.txt', function (error, stdout, stderr) {

                if (error !== null) {
                    console.log('exec error: ' + error);
                }

                // Step 3: Trigger Outlook.exe and Narrator.exe. After opening these two 
                // apps, we will start reading logs.txt, which will be created after 
                // the Narrator script is on.
                open_apps_outlook = exec('START outlook.exe', function (error, stdout, stderr) {

                    if (error !== null) {
                        console.log('exec error: ' + error);
                    }

                    // Open the narrator after Outlook
                    // NOTE: I want the log file to be there once the Narrator is launched
                    open_apps_narrator = exec('START narrator.exe', function (error, stdout, stderr) {

                        if (error !== null) {
                            console.log('exec error: ' + error);
                        }


                        else {
                            // Wait for the logs.txt file to appear
                            openLogFile();
                        }

                    });

                });


            });


        }

    });

