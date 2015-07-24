/// <reference path="uia.d.ts" />
"use strict";


////////////////////////////////////////////////////////////////////// HELPER FUNCTIONS //////////////////////////////////////////////////////////////////////


function testMovingToYesterday() {


}


function testStartReadingOutSuggestions() {

    startReadingSuggestedTimes();

    host.setTimeout(function () {

        ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
        // 1. Don't know how to test the correctness of reading out suggesions yet --> It's almost impossible to make it stable. 


        // Try to move up --> There is no previous suggestions
        suggestedTimesListMoveUp();

        host.setTimeout(function () {

            // Get the sentence that narrator says
            var narratorSentence;

            // For QUnit testing
            QUnit.test("Test for reading out --> no previous suggestions when user presses UP key", function (assert) {

                // Test the output has all the necessary components
                assert.strictEqual(narratorSentence, "You have no previous suggestions.");

            });

            ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // 2. Don't know how to test for reading out --> You have no NEXT suggestions

            testMovingToYesterday();

        }, 5000);

    }, 10000);
}


function testClickingOnSAButton() {

    // Call the function for clicking on SA button
    clickOnSAButton();

    // Test 1: Test if the Narrator is checking whether user has pressed 8 or 
    // not for reading out suggestions
    // - 1. Press UP directly
    suggestedTimesListMoveUp();

    // Wait till the operation to finish
    host.setTimeout(function () {
        ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
        // Get the sentence that narrator says
        var narratorSentence;

        // For QUnit testing
        QUnit.test("Test for if the Narrator is checking whether user has pressed 8 or not for reading out suggestions - Press UP", function (assert) {

            // Test the output has all the necessary components
            assert.strictEqual(narratorSentence, "Make sure you have pressed 8 to start reading suggested times. ");

        });

        // - 2. Press DOWN directly
        suggestedTimesListMoveDown();
        
        host.setTimeout(function () {

            ///////////////////////////////////////////////////////////// TODO /////////////////////////////////////////////////////////////
            // Get the sentence that narrator says
            var narratorSentence;

            // For QUnit testing
            QUnit.test("Test for if the Narrator is checking whether user has pressed 8 or not for reading out suggestions - Press DOWN", function (assert) {

                // Test the output has all the necessary components
                assert.strictEqual(narratorSentence, "Make sure you have pressed 8 to start reading suggested times. ");

            });

            // Then we start the test for start reading out suggestions
            testStartReadingOutSuggestions();

        }, 2000);

    }, 2000);

}

function clickOnNewMeetingButton() {

    return Q.fcall(function () {

        return uia.root().findFirst(function (el) { ((el.name === "New Meeting") && (el.className === "NetUIRibbonButton")); }, 10, 11);

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
        host.setTimeout(function () { clickOnNewMeetingButton(); }, 2000);

        return "Fulfilled";

    }, function (error) { throw new Error("The promise for finding the invoke pattern of calendar button fails. "); });

}


////////////////////////////////////////////////////////////////////// CALL THE TEST FUNCTION //////////////////////////////////////////////////////////////////////
testForSchedulingMeetingScenario();




