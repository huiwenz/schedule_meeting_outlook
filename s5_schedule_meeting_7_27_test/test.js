/// <reference path="uia.d.ts" />
"use strict";

////////////////////////////////////////////////////////////////////// HELPER FUNCTIONS //////////////////////////////////////////////////////////////////////

function jumpToSpecifiedDate(dateStr) {

    return Q.fcall(function () {

        return uia.root().findFirst(function (el) { return ((el.name === "Meeting start date") && (el.className === "RichEdit20WPT")); }, 4, 5);

    }).then(function (meetingStartDateArea) {

        // Get value pattern
        return meetingStartDateArea.getPattern(10002);

    }, function (error) { throw new Error("The promise for finding the meeting start date text area fails. "); })

    .then(function (meetingStartDateAreaValuePat) {

        meetingStartDateAreaValuePat.setValue(dateStr);

        // Randomly change the focus in order to make the date to be set
        // Set focus on the end date 
        return uia.root().findFirst(function (el) { return ((el.name === "Meeting end date") && (el.className === "RichEdit20WPT")); }, 4, 5);

    }, function (error) { throw new Error("The promise for finding the value pattern of meeting start date text area fails. "); })

    .then(function (meetingEndDateField) {

        meetingEndDateField.setFocus(); // Set focus to the end time to refresh the new start/end time
        return "Fulfilled";

    }, function (error) { throw new Error("The promise for finding the meeting end date UIA element fails. "); });
}


function testAddRooms() {

    addRoomsWrapper();

    host.setTimeout(function () {

        // Check if the window has popped up
        return Q.fcall(function () {
            return uia.root().findFirst(function (el) { return (el.name === "Select Rooms: All Rooms"); }, 1, 2);
        }).then(function (addRoomWindow) {

            assert.ok((addRoomWindow !== null), "FROM TEST.JS: TEST IF SELECT ROOMS WINDOW HAS POPPED UP - RESULT OF ASSERTION");

        }, function (error) { throw new Error("The promise for finding adding new rooms window fails. "); });


    }, 3000);

}

// Private function: Takes in a date string, and see if the value of meeting start date text box matches
// the input text string
function checkForDate(matchDate, testMsg, nextStepCallback) {

    // Make sure that we are in Jun 8th
    // Check the UIA element
                    
    return Q.fcall(function () {

        return uia.root().findFirst(function (el) { return ((el.name === "Meeting start date") && (el.className === "RichEdit20WPT")); }, 4, 5);

    }).then(function (meetingStartDateArea) {

        // Get value pattern
        return meetingStartDateArea.getPattern(10002);

    }, function (error) { throw new Error("The promise for finding the meeting start date text area fails. "); })
    .then(function (meetingStartDateAreaValuePat) {
        
        var curDateText = meetingStartDateAreaValuePat.getValue();

        // For unit testing
        console.log("DATEAEEEE " + curDateText);
        assertOk((curDateText.indexOf(matchDate) > -1), testMsg);

        // Continue to do next step if we have one
        if (nextStepCallback !== null) {
            nextStepCallback();
        }

    }, function (error) { throw new Error("The promise for finding the value pattern of meeting start date text area fails. "); });
}

function testMovingToYesterday() {


    // If we only check one of the following cases...
    jumpToSpecifiedDate("6/15/2015");

    host.setTimeout(function () {

        jump_in_calendar(false);

        host.setTimeout(checkForDate("6/12/2015",
            "Test for moving to yesterday --> When input day is normal weekday and yesterday is a weekend"
            , function () {

                return 0;

            }), 10000);

    }, 3000);





    //// Test 1: Jump to a weekend, and the function should not accept this
    //// as input
    //jumpToSpecifiedDate("6/7/2015");

    //host.setTimeout(function () {

    //    jump_in_calendar(false);

    //    host.setTimeout(function () {

    //        //////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
    //        // Get the sentence that narrator says
    //        var narratorSentence;

    //        // Unit testing
    //        assertStrictEqual(narratorSentence, "You've selected a weekend. Please select a weekday.",
    //            "Test for moving to yesterday --> When the input day is a weekend");

    //        // Test 2: Jump to a normal day, and the previous day is not a weekend
    //        jumpToSpecifiedDate("6/4/2015");

    //        host.setTimeout(function () {

    //            jump_in_calendar(false);

    //            host.setTimeout(function () {

    //                // Check if the date is correct
    //                checkForDate("6/3/2015",
    //                    "Test for moving to yesterday --> When input day is normal weekday and yesterday is not weekend",
    //                    function () {

    //                        // Test 2: Jump to a normal day, and the previous day is a weekend
    //                        jumpToSpecifiedDate("6/15/2015");

    //                        host.setTimeout(function () {

    //                            jump_in_calendar(false);

    //                            host.setTimeout(checkForDate("6/12/2015",
    //                                "Test for moving to yesterday --> When input day is normal weekday and yesterday is a weekend"
    //                                , function () {

    //                                    // Test 3: Jump to a day which previou weekday is not visible
    //                                    jumpToSpecifiedDate("6/1/2015");

    //                                    host.setTimeout(function () {

    //                                        jump_in_calendar(false);

    //                                        host.setTimeout(function () {
    //                                            checkForDate("5/29/2015",
    //                                                "Test for moving to yesterday --> When input day is normal weekday and yesterday is in previous month",
    //                                                testAddRooms);
    //                                        }, 6000);

    //                                    }, 5000);

    //                                }), 7000);

    //                        }, 5000);

    //                    });

    //            }, 5000);

    //        }, 5000);

    //    }, 5000);

    //}, 5000);

}

// The helper function for moving back to yesterday
function testMovingToTomorrow() {

    // Test 1: Jump to a weekend, and the function should not accept this
    // as input

    jumpToSpecifiedDate("6/7/2015");

    host.setTimeout(function () {
        
        // Before jump in calendar, let's set up the necessary global variable inside of app.js
        jump_in_calendar(true);

        host.setTimeout(function () {
            //////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // Get the sentence that narrator says
            var narratorSentence;

            // Unit testing
            assertStrictEqual(narratorSentence, "You've selected a weekend. Please select a weekday.",
                "Test for moving to tomorrow --> When the input day is a weekend passed. ");


            // Test 2: Jump to tomorrow from a normal day
            jumpToSpecifiedDate("6/8/2015");

            host.setTimeout(function () {
                jump_in_calendar(true);

                host.setTimeout(function () {
                    checkForDate("6/9/2015",
                                "Test for moving to tomorrow --> When the input day is a normal day and next day is not weekend passed! ",
                                function () {

                                    // Test 3: Tomorrow is a weekend
                                    jumpToSpecifiedDate("6/12/2015");

                                    host.setTimeout(function () {

                                        jump_in_calendar(true);
                                        host.setTimeout(function () {

                                            checkForDate("6/15/2015",
                                                "Test for moving to tomorrow --> When the input day is a normal day and next day is weekend passed! ", testMovingToYesterday);
                                        }, 5000);

                                    }, 3000);

                                });
                }, 5000);

            }, 3000);

        }, 3000);

        console.log("Done");

    }, 3000);

}


function testStartReadingOutSuggestions() {

    startReadingSuggestedTimes();

    host.setTimeout(function () {

        ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
        // 1. Don't know how to test the correctness of reading out suggesions yet --> It's almost impossible to make it stable. 

        // Try to move up --> There is no previous suggestions
        suggestedTimesListMoveUp();

        host.setTimeout(function () {

            //////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // Get the sentence that narrator says
            var narratorSentence;

            assertStrictEqual(narratorSentence, "You have no previous suggestions.",
                "Test for reading out --> no previous suggestions when user presses UP key passed!");

            ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // 2. Don't know how to test for reading out --> You have no NEXT suggestions

            testMovingToYesterday();

        }, 5000);

    }, 10000);
}

// The test steps after clicking on Scheduling Assistant button
function testClickingOnSAButton() {

    // Call the function for clicking on SA button
    clickOnSAButton();

    // Test 1: Test if the Narrator is checking whether user has pressed 8 or 
    // not for reading out suggestions
    // - 1. Press UP directly

    ///////////////////////////////////////////////////////////// UNCOMMENT IT /////////////////////////////////////////////////////////////
    // suggestedTimesListMoveUp();

    // Wait till the operation to finish
    host.setTimeout(function () {
        ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
        // Get the sentence that narrator says
        var narratorSentence;

        assertStrictEqual(narratorSentence, "Make sure you have pressed 8 to start reading suggested times. ",
            "Test for if the Narrator is checking whether user has pressed 8 or not for reading out suggestions - Press UP passed! ");

        // - 2. Press DOWN directly
        suggestedTimesListMoveDown();
        
        host.setTimeout(function () {

            ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // Get the sentence that narrator says
            var narratorSentence;

            assertStrictEqual(narratorSentence, "Make sure you have pressed 8 to start reading suggested times. ",
                "Test for if the Narrator is checking whether user has pressed 8 or not for reading out suggestions - Press DOWN passed! ");

            // Then we start the test for start reading out suggestions
            testStartReadingOutSuggestions();

        }, 2000);

    }, 2000);

}

function clickOnNewMeetingButton() {

    return Q.fcall(function () {

        return uia.root().findFirst(function (el) { return ((el.name === "New Meeting") && (el.className === "NetUIRibbonButton")); }, 10, 11);

    }).then(function (newMeetingButton) {

        return newMeetingButton.getPattern(10000);

    }, function (error) { throw new Error("The promise for finding the New Meeting button fails.  "); })
    .then(function (newMeetingButtonInvokePat) {

        newMeetingButtonInvokePat.invoke();

        // Start to test the function for opening Scheduling Assistant window
        host.setTimeout(function () { testClickingOnSAButton(); }, 3000);

    }, function (error) { throw new Error("The promise for finding the invoke pattern of new meeting pattern fails. "); });

}


////////////////////////////////////////////////////////////////////// MAIN TEST FUNCTION //////////////////////////////////////////////////////////////////////
function testForSchedulingMeetingScenario() {

    // Get into calendar tab by clicking on the Calendar button
    return Q.fcall(function () {
        return uia.root().findFirst(function (el) { return (el.name === "Calendar"); }, 0, 3);
    }).then(function (calendarButton) {

        // Get invoke Pattern
        return calendarButton.getPattern(10000);

    }, function (error) { throw new Error("The promise for finding the calendar button fails."); })

    .then(function (calendarInvoke) {

        // Click on the Calendar tab button
        calendarInvoke.invoke();

        // Wait 2 seconds till we jump to the Calendar page
        // and do the subsequent operations
        host.setTimeout(function () { clickOnNewMeetingButton(); }, 3000);

        return "Fulfilled";

    }, function (error) { throw new Error("The promise for finding the invoke pattern of calendar button fails. "); });

}


////////////////////////////////////////////////////////////////////// CALL THE TEST FUNCTION //////////////////////////////////////////////////////////////////////
// testForSchedulingMeetingScenario();
testMovingToYesterday();



