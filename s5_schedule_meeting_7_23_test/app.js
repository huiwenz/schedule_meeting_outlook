/// <reference path="uia.d.ts" />
"use strict";

/////////////////////////////////////////////////// Global Variable ///////////////////////////////////////////////////
// Wrap everything inside one object to reduce the use of global variable
var scheduleMeetingScript = {
    cur_suggestion_item: null,
    is_activateKey_pressed: false,
    start_reading_suggestions: false,
    cur_date_num: null
}

host.onActivate = function () {
    console.log("onActivate");
};

host.onDeactivate = function () {
    console.log("onDeactivate");
};

host.onSetFocus = function () {
    console.log("onSetFocus");

    var el = uia.focused();

    console.log("name: \"" + el.name + "\"");
    console.log("class name: \"" + el.className + "\"");
    console.log("id: " + el.id);
    console.log("automationid: " + el.automationid);

};

host.onKillFocus = function () {
    console.log("onKillFocus");
}

///////////////////////////////////////////////////// Helper functions /////////////////////////////////////////////////////

// A helper function for changing the date -- Given the time string
function changeTheDate(time) {

    return Q.fcall(function () {

        return uia.root().findFirst(function (el) { return (el.name === "Meeting start time"); }, 0, 5);
    }).then(function (meetingStartTimeBox) {

        // return the value pattern
        return meetingStartTimeBox.getPattern(10002);

    }, function (error) { throw new Error("The promise for locating the start time of the meeting fails. "); })
    .then(function (editBoxValuePattern) {

        editBoxValuePattern.setValue(time);

        // Set focus on the end time to finalize the action
        return uia.root().findFirst(function (el) { return (el.name === "Meeting end time"); }, 0, 5);

    }, function (error) { throw new Error("The promise for getting the value pattern of meeting start time edit box fails. "); })
    .then(function (endTimeEditBox) {

        endTimeEditBox.setFocus(); // Set focus to the end time to refresh the new start/end time
        return "Fulfilled";

    }, function (error) { throw new Error("The promise for finding the end time edit box fails. "); });

}


// continueRoomSelection: A function which contains all the subsequent moves after clicking Add rooms... button
function continueRoomSelection() {

    return Q.fcall(function () {
        return uia.root().findFirst(function (el) { return (el.name.indexOf("Select Rooms") > -1); }, 0, 2);
    }).then(function (selRoomWindow) {

        // Check More columns radio button
        return Q.fcall(function () {
            return selRoomWindow.findFirst(function (el) { return (el.name === "More columns"); }, 0, 1);
        }).then(function (moreColumnsButton) {

            // Click on the button
            var invokePattern = moreColumnsButton.getPattern(10000);
            invokePattern.invoke();

            // One possible question: Do we have to wait?
            // Locate "Search:" edit box
            return selRoomWindow.findFirst(function (el) { return ((el.name === "Search:") && (el.className === "RichEdit20W")); }, 0, 1);

        }, function (error) { throw new Error("The promise for locating More columns radio button fails."); })
        .then(function (editBox) {

            // Begin editing - Type the desired room number
            // Set the focus on edit box
            editBox.setFocus();

            // And users can do all the subsequent steps by themselves
            ////////////////////////////////////////// EATEN //////////////////////////////////////////
            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
            narrator.say("Type the building number or any keyword about the room. After you finish, press enter to start searching.");

            return "Fulfilled";

        }, function (error) { throw new Error("The promise for finding edit box fails."); });

    }, function (error) { throw new Error("The promise for finding the Select Rooms... window fails. "); });

}

// read_out_suggestion: Helps going through and reading out suggestions
function read_out_suggestion(item) {

    var finalReadResult = item.name + ". "; // A sentence, so we need to add a period.

    if (item.name.indexOf("No available rooms") <= -1) { // There is available rooms

        if (item.name.indexOf("conflict") > -1) { // There is a conflict

            // Then we read out the attendees who have a conflict
            // Find the window

            return Q.fcall(function () {
                return uia.root().findFirst(function (el) {
                    return ((el.name.indexOf("All Attendees Status") > -1) && (el.className === "AfxWndW"));
                }, 0, 5);
            }).then(function (attendeeStatusWindow) {

                var attendeeStatus = attendeeStatusWindow.name;

                var statusArray = attendeeStatus.split(';');

                // Record all the BUSY status
                for (var i = 2; i < statusArray.length; i++) {

                    if (statusArray[i].indexOf("Free") <= -1) { // Not free = conflict

                        // Add this status to the final result
                        finalReadResult += (statusArray[i] + ". ");
                    }
                }

                // Read all of them out
                ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
                narrator.say(finalReadResult);
                console.log("Narrator should say: " + finalReadResult);

                return "Fulfilled";

            }, function (error) { throw new Error("Can't find attendees status window!"); });

        } else {

            // Just read out the name of the suggestion
            ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
            narrator.say(finalReadResult);
            console.log("Narrator should say: " + finalReadResult);

        }

    } else {

        // Just read out the name of the suggestion
        ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
        narrator.say(finalReadResult);
        console.log("Narrator should say: " + finalReadResult);
    }

}


///////////////////////////////////// Helper function for switching back and forth between today and tomorrow /////////////////////////////////////
// jump_in_calendar - Given a boolean "forward" to the function to tell it to jump to tomorrow or to yesterday. When forward = true,
// we jump to tomorrow, and to yesterday when it is set to false. 
function jump_in_calendar(forward) {

    // Locate the Calendar Control on the top right
    return Q.fcall(function () { return uia.root().findFirst(function (el) { return (el.name === "Calendar Control"); }, 0, 11); })

    .then(function (calendarControlWindow) {

        // Find the date that is currently focused on Calendar Control
        // return calendarControlWindow.lastChild(); Still doesn't work.
        return calendarControlWindow.childNodes()[2];

    }, function (error) { throw new Error("The promise for finding Calendar Control window fails."); })
    .then(function (curDate) {
        var curDateStr = curDate.name;

        // Parse out the date number
        scheduleMeetingScript.cur_date_num = ((curDateStr.split(",")[1]).split(" "))[2];

        // Locate the dates table
        // return curDate.lastChild(); Still doesn't work
        return curDate.childNodes()[1];

    }, function (error) { throw new Error("The promise for finding current date element fails."); })
    .then(function (datestable) {

        // first week
        return datestable.firstChild().nextSibling();

    }, function (error) { throw new error("The promise for finding the dates table fails. "); })
    .then(function (firstWeekCellGroup) {

        console.log("Find first week cell group");
        // Locate the first one
        // Naive loop
        var first_cell = firstWeekCellGroup.firstChild();

        while (first_cell !== null) {

            if (first_cell.name[1] === '1') {
                break;
            }
            // Go through all the cells
            first_cell = first_cell.nextSibling();

        }

        // After finding the first cell of the month, we want to find the current day which has the name as curDateStr
        // NOTE: cur_cell must be a WEEKDAY. 

        // Keep finding
        while (first_cell.name !== scheduleMeetingScript.cur_date_num) {

            if (first_cell.nextSibling() !== null) {

                first_cell = first_cell.nextSibling();

            } else {

                // Reached the end of a week
                // Find the parent
                var thisWeek = first_cell.parentNode();

                var nextWeek = thisWeek.nextSibling(); // This will not be null

                first_cell = nextWeek.firstChild();

                // Handling errors
                if (first_cell == null) {
                    throw new Error("Unable to locate the next date cell in next week. Something is wrong.");
                }

            }

        }

        // Until we find it
        // Check if it is a WEEKEND
        if (first_cell.nextSibling() == null || first_cell.previousSibling() == null) { // it is a weekend
            ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
            narrator.say("You've selected a weekend. Please select a weekday.");
        } else {

            if (forward) {
                find_next_sibling(first_cell);
            } else {
                find_prev_sibling(first_cell);
            }
            /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
            // Current implementation doesn't keep track of current day. This is different from the spec. 
            /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
        }

        return "Fulfilled";

    }, function (error) { throw new Error("The promise for finding cells for the first week fails."); });

}

// Find_prev_sibling - Given a date cell, c [which is a weekday] in the Calendar Control,
// click on its previous sibling which is not a weekend.
function find_prev_sibling(dateCell) { 


    return Q.fcall(function () {

        return dateCell.previousSibling();

    }, function (error) { throw new Error("The promise for finding the previous sibling of input cell fails. "); })

    .then(function (dateCellPrevious) {

        if (dateCellPrevious != null) {

            return Q.fcall(function () {
                return dateCellPrevious.previousSibling();
            }).then(function (dateCellPreviousPrev) {


                if (dateCellPreviousPrev != null) {

                    return Q.fcall(function () {

                        return dateCellPrevious.nextSibling();

                    }).then(function (dateCellPreviousNext) {

                        if (dateCellPreviousNext != null) {

                            return Q.fcall(function () { return dateCellPrevious.getPattern(10000); })
                            .then(function (dateCellPreviousInvokePattern) {

                                dateCellPreviousInvokePattern.invoke();

                            }, function (error) { throw new Error("The promise for getting the invoke pattern of previous date cell fails."); });

                        } else {
                            ////////////////////////////////////////// EATEN //////////////////////////////////////////
                            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
                            narrator.say("Something is wrong. The previous day of a weekday can't be a Saturday. Please enter a valid weekday. ");
                        }

                        return "Fulfilled";

                    }, function (error) { throw new Error("The promise for finding the next sibling of previous date cell fails. ") });

                } else { // is NULL

                    // Then we go back to find last Friday

                    return Q.fcall(function () {
                        return dateCellPrevious.parentNode();
                    }).then(function (dateCellWeek) {

                        return dateCellWeek.previousSibling();

                    }, function (error) { throw new Error("The promise for finding the week cell group of previous date cell fails. "); })
                    .then(function (dateCellPrevWeek) {

                        return dateCellPrevWeek.firstChild(); // Get the first child of the previous week

                    }, function (error) { throw new Error("The promise for finding the previous week of current week cell group fails. "); })

                    .then(function (firstDayPrevWeek) {

                        // Check if current week is the first week of the month
                        if (firstDayPrevWeek != null) {

                            if (firstDayPrevWeek.name === "Su") { // We need to switch back to previous month

                                // Press the Previous button to go back to the previous month
                                // We don't return a promise for now

                                return Q.fcall(function () {
                                    return uia.root().findFirst(function (el) {
                                        return ((el.name === "Previous Button") && (el.parentNode().name === "Calendar Control")); // Still need to check if this is the best way
                                    }, 1, 11);
                                }).then(function (previousButton) {

                                    var invokePatternPrevButton = previousButton.getPattern(10000);
                                    invokePatternPrevButton.invoke();

                                    /////////////////////////////////////////////////////////
                                    // After invoking, you need to wait for 2 seconds
                                    host.setTimeout(function () {

                                        return Q.fcall(function () {
                                            // Relocate Calendar Control window
                                            // Locate the Calendar Control on the top right
                                            return uia.root().findFirst(function (el) { return (el.name === "Calendar Control"); }, 0, 11);
                                        }).then(function (calendarControlWindow) {

                                            ////////////////////////////////////////////// TO BE CHANGED  ////////////////////////////////////////////// 
                                            // Change to lastChild() later
                                            return calendarControlWindow.firstChild().nextSibling().nextSibling();


                                        }, function (error) { throw new Error("The promise for finding the Calendar Control window fails. ") })
                                        .then(function (curDate) {

                                            // Find the dates table
                                            return curDate.childNodes()[1];

                                        }, function (error) { throw new Error("The promise for finding current date element fails. "); })

                                        .then(function (datesTable) {

                                            return datesTable.childNodes()[5];

                                        }, function (error) { throw new Error("The promise for finding the dates table fails. "); })

                                        .then(function (secondLastWeek) {

                                            // Locate a Friday
                                            return secondLastWeek.childNodes()[5];

                                        }, function (error) { throw new Error("The promise for finding the second last week fails. ") })

                                        .then(function (targetDay) {

                                            var prevWeekdayInvokePattern = targetDay.getPattern(10000);

                                            prevWeekdayInvokePattern.invoke(); // Click on it

                                            return "Fulfilled";

                                        }, function (error) { throw new Error("The promise for finding the Friday in the second last week fails. "); });

                                    }, 2000);

                                    return "Fulfilled";

                                }, function (error) { throw new Error("The promise for finding the previous button on Calendar Control fails. "); });


                            } else { // Just find the 6th child of the previous week -- A Friday


                                return Q.fcall(function () {
                                    return dateCellPrevious.parentNode();
                                }).then(function (dateCellWeek) {

                                    return dateCellWeek.previousSibling();

                                }, function (error) { throw new Error("The promise for finding the week cell group of previous date cell fails. "); })

                                .then(function (dateCellPrevWeek) {

                                    return dateCellPrevWeek.childNodes()[5];
                                }, function (error) { throw new Error("The promise for finding the previous week of current date cell fails. "); })

                                .then(function (prevTargetDay) {

                                    var prevWeekdayInvokePattern = prevTargetDay.getPattern(10000);

                                    prevWeekdayInvokePattern.invoke(); // Click on it

                                }, function (error) { throw new Error("The promise for finding the Friday of last week fails. "); });

                            }

                        } else {
                            ////////////////////////////////////////// might be EATEN //////////////////////////////////////////
                            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
                            narrator.say("Something is wrong. Please do the operation again. ");

                        }

                        return "Fulfilled";

                    }, function (error) { throw new Error("The promise for finding the first child of previous week fails. "); });

                }


            }, function (error) { throw new Error("The promise for finding the previous sibling of previous date cell fails. "); });

        } else {
            
            return "Fulfilled";
            throw new Error("Error. The cell might not be a weekday. "); // Error
        }


    }, function (error) { throw new Error("The promise for checking the previous sibling fails. "); });

}


// Find_next_sibling - Given a date cell, c [which is a weekday] in the Calendar Control,
// click on its next sibling which is not a weekend.
function find_next_sibling(dateCell) {

    return Q.fcall(function () {

        return dateCell.nextSibling();

    }).then(function (nextOne) {

        if (nextOne != null) {

            return Q.fcall(function () {

                return nextOne.previousSibling();

            }).then(function (previousSib) {

                if (previousSib != null) {

                    return Q.fcall(function () {
                        return nextOne.nextSibling();
                    }).then(function (nextSib) {

                        if (nextSib == null) {

                            return Q.fcall(function () {
                                return nextOne.parentNode();
                            }).then(function (thisWeek) {

                                return thisWeek.nextSibling();

                            }, function (error) { throw new Error("The promise for finding the week cell group of next cell fails. "); })

                            .then(function (nextWeek) {

                                return nextWeek.childNodes()[1];

                            }, function (error) { throw new Error("The promise for finding the next week fails. "); })

                            .then(function (firstDayNextWeek) {

                                return firstDayNextWeek.getPattern(10000);

                            }, function (error) { throw new Error("The promise for finding the first day of next week fails. "); })

                            .then(function (firstDayInvokePat) {

                                firstDayInvokePat.invoke();
                                return "Fulfilled";

                            }, function (error) { throw new Error("The promise for finding the invoke pattern for first day on next week fails. "); });

                        } else {
                            return Q.fcall(function () {

                                return nextOne.getPattern(10000);

                            }).then(function (nextOneInvokePattern) {

                                nextOneInvokePattern.invoke();
                                return "Fulfilled";

                            }, function (error) { throw new Error("The promise for finding teh invok pattern for next day fails."); });
                        }

                    }, function (error) { throw new Error("The promise for finding the next sibling of next day fails. "); });

                } else {
                    throw new Error("There is something wrong. Previous sibling can't be null. ");
                }

            }, function (error) { throw new Error("The promise for finding the previous sibling of next day fails. "); });

        } else {
            // The input day is not a weekday
            ////////////////////////////////////////// MIGHT BE EATEN //////////////////////////////////////////
            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
            narrator.say("The input day is not a weekday. Please enter a weekday. ");
        }

    }, function (error) { throw new Error("The promise for finding the next sibling of date cell fails. "); });
}


// The wrapper function for clicking on Scheduling Assistant window
function clickOnSAButton() {

    // Reset the activate key
    if (scheduleMeetingScript.is_activateKey_pressed) {
        scheduleMeetingScript.is_activateKey_pressed = false;
    }

    // Click on "Scheduling Assistant" button
    return Q.fcall(function () {
        return uia.root().findFirst(function (el) { return (el.name === "Scheduling Assistant"); }, 0, 11);
    }).then(function (schedulingAssistantButton) {


        var togglePat = schedulingAssistantButton.getPattern(10015);
        togglePat.toggle();

        // inform the user
        ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
        narrator.say("You are in the scheduling assistant tab. Please choose attendees and room. ");
        return "fulfilled";

    }, function (error) { throw new Error("The promise of finding Scheduling Assistant tab item fails. "); });

}


function suggestedTimesListMoveDown() {

    if (scheduleMeetingScript.is_activateKey_pressed) {

        if (scheduleMeetingScript.start_reading_suggestions && (scheduleMeetingScript.cur_suggestion_item !== null)) {

            if (scheduleMeetingScript.cur_suggestion_item.nextSibling() !== null) { // There is a next item
                // Note: We need to consider the case when previousSibling is the vertical scroll bar

                scheduleMeetingScript.cur_suggestion_item = scheduleMeetingScript.cur_suggestion_item.nextSibling(); // Find the next sibling

                // Instead, we use an another way -- Type the time into the Start time and end time region
                // Just type the start time would suffice
                var nameStr = scheduleMeetingScript.cur_suggestion_item.name;

                // Locate the text box for typing Meeting Start time
                // And change the time - Only enter the start time
                changeTheDate(nameStr.split("-")[0]);

                // Wait for 3 seconds
                host.setTimeout(function () { read_out_suggestion(scheduleMeetingScript.cur_suggestion_item); }, 2000);

            } else {
                ////////////////////////////////////////// EATEN //////////////////////////////////////////
                ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
                narrator.say("You have no next suggestions.");

            }



        } else {
            ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
            narrator.say("Make sure you have pressed 8 to start reading suggested times. ");
        }

        scheduleMeetingScript.is_activateKey_pressed = false;
    }

}

function suggestedTimesListMoveUp() {

    if (scheduleMeetingScript.is_activateKey_pressed) {

        if (scheduleMeetingScript.start_reading_suggestions && (scheduleMeetingScript.cur_suggestion_item !== null)) {

            if ((scheduleMeetingScript.cur_suggestion_item.previousSibling() !== null) && (scheduleMeetingScript.cur_suggestion_item.previousSibling().name !== "Vertical")) { // There is a previous item
                // Note: We need to consider the case when previousSibling is the vertical scroll bar

                scheduleMeetingScript.cur_suggestion_item = scheduleMeetingScript.cur_suggestion_item.previousSibling(); // Find the previous sibling

                // Instead, we use an another way -- Type the time into the Start time and end time region
                // Just type the start time would suffice
                var nameStr = scheduleMeetingScript.cur_suggestion_item.name;

                // Locate the text box for typing Meeting Start time
                // And change the time - Only enter the start time
                changeTheDate(nameStr.split("-")[0]);

                // Wait for 3 seconds
                host.setTimeout(function () { read_out_suggestion(scheduleMeetingScript.cur_suggestion_item); }, 2000);

            } else {
                ////////////////////////////////////////// MIGHT BE EATEN //////////////////////////////////////////
                ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
                narrator.say("You have no previous suggestions.");
            }


        } else {

            ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
            narrator.say("Make sure you have pressed 8 to start reading suggested times. ");

        }

        scheduleMeetingScript.is_activateKey_pressed = false;
    }

}


function startReadingSuggestedTimes() {

    // Reset the activate key
    if (scheduleMeetingScript.is_activateKey_pressed) {
        scheduleMeetingScript.is_activateKey_pressed = false;
    }

    scheduleMeetingScript.start_reading_suggestions = true;

    return Q.fcall(function () {
        return uia.root().findFirst(function (el) { return ((el.name === "MsoDockRight") && (el.className === "MsoCommandBarDock")); }, 0, 2);
    })

   .then(function (msoDockRightWindow) {

       // Locate suggested time list
       return msoDockRightWindow.findFirst(function (el) { return (el.name === "Suggested times:"); }, 0, 8);

   }, function (error) { throw new Error("The promise for finding msoDockRightWindow fails."); })

    .then(function (suggestedTimesText) {

        // This is just the text "Suggested times:"
        // The list is actually its next sibling
        return suggestedTimesText.nextSibling();

    }, function (error) { throw new Error("The promise for finding suggested times text fails."); })

    .then(function (suggestedTimesList) {

        // In the list, we first locate on the first item
        // Add error handling in here
        ///////////////////////////////////////////////// THERE ARE TWO CASES: WITH SCROLL BAR AND WITHOUT SCROLL BAR /////////////////////////////////////////////////
        scheduleMeetingScript.cur_suggestion_item = suggestedTimesList.firstChild();
        if (scheduleMeetingScript.cur_suggestion_item.name === "Vertical") { // If the first child is the scroll bar??
            // Should define to the next sibling
            scheduleMeetingScript.cur_suggestion_item = suggestedTimesList.firstChild().nextSibling();
        }


        // Reading out the first child
        if (scheduleMeetingScript.cur_suggestion_item !== null) {

            // Click on the item
            var invokePattern = scheduleMeetingScript.cur_suggestion_item.getPattern(10000);
            invokePattern.invoke();

            // Wait for 2 seconds
            host.setTimeout(function () { read_out_suggestion(scheduleMeetingScript.cur_suggestion_item); }, 2000);


        } else {
            ////////////////////////////////////////// MIGHT BE EATEN //////////////////////////////////////////
            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
            narrator.say("There is no suggested times.");
        }

        return "fulfilled";

    }, function (error) { throw new Error("The promise for finding suggested times list fails. "); });

}


//////////////////////////////////////////// Host.onKeypress ////////////////////////////////////////////
host.onKeypress = function (e) {
    console.log("onkeypress");
    console.log(JSON.stringify(e));

    // "1"
    // Act as the "Activation" key
    if (e.keyCode === 49) {
        
        scheduleMeetingScript.is_activateKey_pressed = true;
    }

    // "2"
    // When "Activate" key is pressed, it will serve as the key that can help user
    // to move around suggested times - UP
    else if (e.keyCode === 50) {

        suggestedTimesListMoveUp();

    }


    // "3"
    // When "Activate" key is pressed, it will serve as the key that can help user
    // to move around suggested times - DOWN
    else if (e.keyCode === 51) {
        suggestedTimesListMoveDown();
    }

    // "4"
    else if (e.keyCode === 52) {
        // Reset the activate key
        if (scheduleMeetingScript.is_activateKey_pressed) {
            scheduleMeetingScript.is_activateKey_pressed = false;
        }
    }

    // "5"
    else if (e.keyCode === 53) {

        // Reset the activate key
        if (scheduleMeetingScript.is_activateKey_pressed) {
            scheduleMeetingScript.is_activateKey_pressed = false;
        }

        debugger;
    }

    // "6" - Shortcut key for clicking on "Scheduling Assistant button" to get into the Scheduling
    // Assistant UI
    // Scenario for scheduling a meeting - Assume the user is already on the window of
    // scheduling a new meeting
    else if (e.keyCode === 54) {
        clickOnSAButton();
    }

    // After getting into Scheduling Assistant tab, we can:
    // Press "7" - The shortcut key for adding rooms
    else if (e.keyCode === 55) {

        // Reset the activate key
        if (scheduleMeetingScript.is_activateKey_pressed) {
            scheduleMeetingScript.is_activateKey_pressed = false;
        }

        // Locate "Add rooms..." button
        return Q.fcall(function () {

            return uia.root().findFirst(function (el) { return (el.name.indexOf("Add Rooms") > -1); }, 0, 5);

        }).then(function (addRoomsButton) {

            // Get Invoke pattern
            return addRoomsButton.getPattern(10000);

        }, function (error) {

            ////////////////////////////////////////// MIGHT BE EATEN //////////////////////////////////////////
            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
            narrator.say("Check if you already opened Scheduling Assistant.");
            throw new Error("The promise for opening Scheduling Assistant fails. ");
        }).then(function (addroomButtonInvokePat) {

            addroomButtonInvokePat.invoke();

            // Wait for 2 seconds and continue the room selection
            host.setTimeout(function () { continueRoomSelection(); }, 2000);

            return "fulfilled";

        }, function (error) { throw new Error("The promise for finding the invoke pattern of Add Rooms... button fails. ") });

    }
        // "8"
        // For start reading out suggested times in Scheduling Assistant window
    else if (e.keyCode === 56) {

        startReadingSuggestedTimes();

    }

    // "9"
    // For switching back and forth in the calendar
    else if (e.keyCode === 57) {

        // Reset the activate key
        // The shortcut key for switching back from today to yesterday
        if (scheduleMeetingScript.is_activateKey_pressed) {

            scheduleMeetingScript.is_activateKey_pressed = false;

            jump_in_calendar(false);


        } else {
            // For switching from today to tomorrow
            jump_in_calendar(true);
        }
    }
 
};
