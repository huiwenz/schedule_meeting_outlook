/// <reference path="uia.d.ts" />

host.onActivate = function () {
    console.log("onActivate");
};

host.onDeactivate = function () {
    console.log("onDeactivate");
};

host.onSetFocus = function () {
    //console.log("onSetFocus");

    //var el = uia.focused();

    //console.log("name: \"" + el.name + "\"");
    //console.log("class name: \"" + el.className + "\"");
    //console.log("id: " + el.id);
    //console.log("automationid: " + el.automationid);

};

host.onKillFocus = function () {
    console.log("onKillFocus");
}

///////////////////////////////////////////////////// Helper functions /////////////////////////////////////////////////////

// continueRoomSelection: A function which contains all the subsequent moves after clicking Add rooms... button
function continueRoomSelection(desktop) {

    return Q.fcall(function () {
        return desktop.findFirst(function (el) { return (el.name.indexOf("Select Rooms") > -1); }, 0, 2);
    }).then(function (selRoomWindow) {

        // Check More columns radio button
        return Q.fcall(function () {
            return selRoomWindow.findFirst(function (el) { return (el.name == "More columns"); }, 0, 1);
        }).then(function (moreColumnsButton) {

            // Click on the button
            var invokePattern = moreColumnsButton.getPattern(10000);
            invokePattern.invoke();

            // One possible question: Do we have to wait?
            // Locate "Search:" edit box
            return selRoomWindow.findFirst(function (el) { return ((el.name == "Search:") && (el.className == "RichEdit20W")); }, 0, 1);

        }, function (error) { throw new Error("The promise for locating More columns radio button fails."); })
        .then(function (editBox) {

            // Begin editing - Type the desired room number
            // Set the focus on edit box
            editBox.setFocus();

            // And users can do all the subsequent steps by themselves
            ////////////////////////////////////////// EATEN //////////////////////////////////////////
            ////////////////////////////////////////// TO BE FIXED LATER //////////////////////////////////////////
            narrator.say("Type the building number or any keyword about the room. After you finish, press enter to start searching.");


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
                    return ((el.name.indexOf("All Attendees Status") > -1) && (el.className == "AfxWndW"));
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
                
                

            }, function (error) { throw new Error("Can't find attendees status window!"); })

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


// Helper function for switching back and forth between today and tomorrow

// Find_prev_sibling - Given a date cell, c [which is a weekday] in the Calendar Control,
// return its previous sibling which is not a weekend.
function find_prev_sibling(dateCell) { 

    // Locate the previous sibling
    var dateCellPrevious = dateCell.previousSibling();

    // Record final result
    var finalResult;

    if (dateCellPrevious != null) {

        if ((dateCellPrevious.previousSibling() != null) && (dateCellPrevious.nextSibling() != null)) { // A weekday

            return dateCellPrevious;

        } else if (dateCellPrevious.previousSibling() == null){ // A Sunday --> That is the only case. Can't be a Saturday. 
        
            // Then we go back to find last Friday

            ///////////////////////////////////////////////////////// CHECK ERRORS(MAYBE CHANGE TO PROMISES?) /////////////////////////////////////////////////////////
            // Get the week
            var dateCellWeek = dateCellPrevious.parentNode();

            // Check if the week is the first week of the month
            var dateCellPrevWeek = dateCellWeek.previousSibling();
            var firstDayPrevWeek = dateCellPrevWeek.firstChild();

            if (firstDayPrevWeek != null) {

                if (firstDayPrevWeek.name == "Su") { // We need to switch back to previous month

                    // Press the Previous button to go back to the previous month
                    // We don't return a promise for now

                    var previousButton = uia.root().findFirst(function (el) {
                        return ((el.name == "Previous Button") && (el.parentNode().name == "Calendar Control")); // Still need to check if this is the best way
                    }, 1, 11);

                    if (previousButton != null) {


                        var invokePatternPrevButton = previousButton.getPattern(10000);
                        invokePatternPrevButton.invoke();

                        /////////////////////////////////////////////////////////
                        // After invoking, you need to wait for 2 seconds
                        host.setTimeout(function () {

                            // Return the last weekday of previous month
                            // Second last week
                            ////////////////////////////////////////////// TBD: ERROR HANDLING //////////////////////////////////////////////
                            var calendarControl = uia.root().findFirst(function (el) { return (el.name == "Calendar Control"); }, 0, 10);

                            var weeksTable = calendarControl.lastChild().firstChild().nextSibling();

                            // Second last week
                            if ((weeksTable != null) && (weeksTable.childNodes().length == 7)) {

                                var secondLastWeek = weeksTable.childNodes()[5];

                                return secondLastWeek.childNodes()[5]; // A Friday

                            } else {
                                throw new Error("You didn't find the correct weeks table!");
                            }
                            ////////////////////////////////////////////// TBD: ERROR HANDLING //////////////////////////////////////////////


                        }, 2000);
                        /////////////////////////////////////////////////////////

                    } else {
                        throw new Error("Can't find Previous Button!!")
                    }

                } else { // Just find the 6th child of the previous week -- A Friday


                    if (firstDayPrevWeek.length == 7) {
                        return firstDayPrevWeek.childNodes()[5];
                    } else {
                        throw new Error("There is something wrong with the previous week!!");
                    }

                }

            }

            ///////////////////////////////////////////////////////// CHECK ERRORS /////////////////////////////////////////////////////////
            
        }

    } else {
        throw new Error("Error. The cell might not be a weekday. "); // Error
    }

}

function find_next_sibling(dateCell) {

    var nextOne = dateCell.nextSibling();

    if (nextOne != null) {
        
        if ((nextOne.previousSibling() != null) && (nextOne.nextSibling() != null)) {
            // A weekday
            return nextOne;
        } else if (nextOne.nextSibling() == null) { // A Saturday

            ///////////////////////////////////////////////////// ERROR HANDLING /////////////////////////////////////////////////////
            var thisWeek = nextOne.parentNode();

            if (thisWeek != null) {

                var nextWeek = thisWeek.nextSibling();

                if (nextWeek != null) {

                    return nextWeek.childNodes()[1];


                } else {
                    throw new Error("Can't find next week! Something is wrong.");
                }

            } else {
                throw new Error("Can't find this week cell group! Somethinh is wrong.");
            }

            ///////////////////////////////////////////////////// ERROR HANDLING /////////////////////////////////////////////////////


        }

    } else {

        throw new Error("Error. The cell might not be a weekday.");

    }


}


//////////////////////////////////////////// Global variable ////////////////////////////////////////////
var cur_suggestion_item; // Global var for reading out suggestion times
var cur_date_num;
var cur_day_cell;
var is_activateKey_pressed = false;
var start_reading_suggestions = false;


//////////////////////////////////////////// Host.onKeypress ////////////////////////////////////////////
host.onKeypress = function (e) {
    console.log("onkeypress");
    console.log(JSON.stringify(e));

    // "1"
    // Act as the "Activation" key
    if (e.keyCode === 49) {
        
        is_activateKey_pressed = true;
    }

    // "2"
    // When "Activate" key is pressed, it will serve as the key that can help user
    // to move around suggested times - UP
    else if (e.keyCode === 50) {

        if (is_activateKey_pressed) {

            if (start_reading_suggestions && (cur_suggestion_item != null)) {

                if ((cur_suggestion_item.previousSibling() != null) && (cur_suggestion_item.previousSibling().name != "Vertical")){ // There is a previous item
                    // Note: We need to consider the case when previousSibling is the vertical scroll bar

                    cur_suggestion_item = cur_suggestion_item.previousSibling(); // Find the previous sibling

                    /////////////////////////////////////// DELAY ///////////////////////////////////////
                    ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
                    // narrator.say(cur_suggestion_item.name); // Read the name
                    // console.log("Narrator say: " + cur_suggestion_item.name);
                    /////////////////////////////////////// DELAY ///////////////////////////////////////

                    // Click on the item
                    var invokePattern = cur_suggestion_item.getPattern(10000);
                    invokePattern.invoke();

                    // Wait for 2 seconds
                    host.setTimeout(function () { read_out_suggestion(cur_suggestion_item); }, 2000);
                    // read_out_suggestion(cur_suggestion_item); // And read it out

                } else {

                    narrator.say("You have no previous suggestions.");

                }


            } else {

                ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
                narrator.say("Make sure you have pressed 8 to start reading suggested times. ");

            }

            is_activateKey_pressed = false;
        }

    }

    // "3"
    // When "Activate" key is pressed, it will serve as the key that can help user
    // to move around suggested times - DOWN
    else if (e.keyCode === 51) {

        if (is_activateKey_pressed) {


            if (start_reading_suggestions && (cur_suggestion_item != null)) {

                if (cur_suggestion_item.nextSibling() != null) { // There is a next item
                    // Note: We need to consider the case when previousSibling is the vertical scroll bar

                    cur_suggestion_item = cur_suggestion_item.nextSibling(); // Find the next sibling

                    /////////////////////////////////////// DELAY ///////////////////////////////////////
                    ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
                    // narrator.say(cur_suggestion_item.name); // Read the name
                    // console.log("Narrator say: " + cur_suggestion_item.name);
                    /////////////////////////////////////// DELAY ///////////////////////////////////////

                    // Click on the item
                    var invokePattern = cur_suggestion_item.getPattern(10000);
                    invokePattern.invoke(); // Problem: Unable to click if the tab is not visible
                    // But, how do I know if I've reached the downmost item in the item bar?
                    

                    // Wait for 2 seconds
                    host.setTimeout(function () { read_out_suggestion(cur_suggestion_item); }, 2000);
                    // read_out_suggestion(cur_suggestion_item); // And read it out

                } else {

                    narrator.say("You have no next suggestions.");

                }



            } else {

                ////////////////////////////////////////////////////////// NOTE: MIGHT BE EATEN //////////////////////////////////////////////////////////
                ///////////////////////////////////////////////////////// TO BE FIXED LATER /////////////////////////////////////////////////////////
                narrator.say("Make sure you have pressed 8 to start reading suggested times. ");

            }

            is_activateKey_pressed = false;
        }

    }

    // "5"
    else if (e.keyCode === 53) {
        debugger;
    }

        // "6" - Shortcut key for clicking on "Scheduling Assistant button" to get into the Scheduling
        // Assistant UI
        // Scenario for scheduling a meeting - Assume the user is already on the window of
        // scheduling a new meeting
    else if (e.keyCode === 54) {

        return Q.fcall(function () {

            // Find desktop element
            return uia.root();

        }).then(function (desktop) {

            // Click on "Scheduling Assistant" button
            return Q.fcall(function () {
                return desktop.findFirst(function (el) { return (el.name == "Scheduling Assistant"); }, 0, 11)
            }).then(function (schedulingAssistantButton) {

                
                var togglePat = schedulingAssistantButton.getPattern(10015);
                togglePat.toggle();

                // inform the user
                // NOTE: Maybe after waiting for 2 seconds?
                narrator.say("You are in the scheduling assistant tab. Please choose attendees and room. ");


            }, function (error) { throw new Error("The promise of finding Scheduling Assistant tab item fails. "); });

        }, function (error) { throw new Error("The promise of finding root element(desktop) fails. "); });

    }

        // After getting into Scheduling Assistant tab, we can:
        // Press "7" - The shortcut key for adding rooms
    else if (e.keyCode === 55) {

        return Q.fcall(function () {
            return uia.root();
        }).then(function (desktop) {

            // Locate "Add rooms..." button
            return Q.fcall(function () {

                return desktop.findFirst(function (el) { return (el.name.indexOf("Add Rooms") > -1); }, 0, 5);

            }).then(function (addRoomsButton) {

                // Get Invoke pattern
                return addRoomsButton.getPattern(10000);

            }, function (error) {
                narrator.say("Check if you already opened Scheduling Assistant.");
                throw new Error("The promise for opening Scheduling Assistant fails. ");
            }).then(function (addroomButtonInvokePat) {

                addroomButtonInvokePat.invoke();

                // Wait for 2 seconds and continue the room selection
                host.setTimeout(function () { continueRoomSelection(desktop); }, 2000);

            }, function (error) { throw new Error("The promise for finding the invoke pattern of Add Rooms... button fails. ") });

        }, function (error) { throw new Error("The promise for finding desktop fails. "); });

    }
        // "8"
        // For reading out suggested times in Scheduling Assistant window
    else if (e.keyCode === 56) {



        // The shortcut key for switching back from today to yesterday
        if (is_activateKey_pressed) {


            is_activateKey_pressed = false;

            ////////////////////////////////////////////////////// CHANGE /////////////////////////////////////////////////
            ///////////////////////////////////////////////// REFACTOR THE STEPS OF FINDING CURRENT CELL INTO A FUNCTION


            // REPLACE UIA.ROOT()

            // Locate the Calendar Control on the top right
            return Q.fcall(function () { return uia.root().findFirst(function (el) { return (el.name == "Calendar Control"); }, 0, 11); })

            .then(function (calendarControlWindow) {

                calendarControlWindow.setFocus();

                // Find the date that is currently focused on Calendar Control
                /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
                // return calendarControlWindow.lastChild(); 
                /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////

                ////////////////////////HARDCODED
                return calendarControlWindow.firstChild().nextSibling().nextSibling();

            }, function (error) { throw new Error("The promise for finding Calendar Control window fails."); })
            .then(function (curDate) {
                var curDateStr = curDate.name;

                // Parse out the date number
                //////////////////////////////////////////////////////// ERROR HANDLING ////////////////////////////////////////////////////////
                cur_date_num = ((curDateStr.split(",")[1]).split(" "))[2];

                // Locate the dates table
                /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
                // return curDate.lastChild();
                /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
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

                while (first_cell != null) {

                    /////////////////////////////////////////////// ERROR CHECKING /////////////////////////////////////////////// 
                    if (first_cell.name[1] === '1') {

                        break;
                    }

                    // Go through all the cells
                    first_cell = first_cell.nextSibling();

                }

                // After finding the first cell of the month, we want to find the current day which has the name as curDateStr
                // NOTE: cur_cell must be a WEEKDAY. 

                // Keep finding
                while (first_cell.name != cur_date_num) {

                    if (first_cell.nextSibling() != null) {

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
                if (first_cell.nextSibling() == null || first_cell.previousSibling() == null) { // it is weekend

                    //////////////////////////////////////////////////// BREAK? How to? ///////////////////////////////////////////////////
                    narrator.say("You've selected a weekend. Please select a weekday.");

                } else {

                    var nextWeekday = find_prev_sibling(first_cell);

                    var nextWeekdayInvokePattern = nextWeekday.getPattern(10000);

                    nextWeekdayInvokePattern.invoke(); // Click on it

                    /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
                    // Current implementation doesn't keep track of current day. This is different from the spec. 
                    /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
                }

            }, function (error) { throw new Error("The promise for finding cells for the first week fails."); });


        } else {

            // For reading out suggested times in Scheduling Assistant window

            start_reading_suggestions = true;

            return Q.fcall(function () {
                return uia.root();
            })

            .then(function (desktop) {

                return desktop.findFirst(function (el) { return ((el.name == "MsoDockRight") && (el.className == "MsoCommandBarDock")); }, 0, 2);

            }, function (error) { throw new Error("The promise for finding the root element fails."); })

           .then(function (msoDockRightWindow) {

               // Locate suggested time list
               return msoDockRightWindow.findFirst(function (el) { return (el.name == "Suggested times:"); }, 0, 8);

           }, function (error) { throw new Error("The promise for finding msoDockRightWindow fails."); })

            .then(function (suggestedTimesText) {

                // This is just the text "Suggested times:"
                // The list is actually its next sibling
                return suggestedTimesText.nextSibling();

            }, function (error) { throw new Error("The promise for finding suggested times text fails."); })

            .then(function (suggestedTimesList) {

                suggestedTimesList.setFocus(); // Just for test

                // In the list, we first locate on the first item
                // Add error handling in here
                ///////////////////////////////////////////////// THERE ARE TWO CASES: WITH SCROLL BAR AND WITHOUT SCROLL BAR /////////////////////////////////////////////////
                cur_suggestion_item = suggestedTimesList.firstChild();
                if (cur_suggestion_item.name == "Vertical") { // If the first child is the scroll bar??
                    // Should define to the next sibling
                    cur_suggestion_item = suggestedTimesList.firstChild().nextSibling();
                }


                // Reading out the first child
                if (cur_suggestion_item != null) {

                    // Click on the item
                    var invokePattern = cur_suggestion_item.getPattern(10000);
                    invokePattern.invoke();

                    // Wait for 2 seconds
                    host.setTimeout(function () { read_out_suggestion(cur_suggestion_item); }, 2000);


                } else {
                    narrator.say("There is no suggested times.");
                }


            }, function (error) { throw new Error("The promise for finding suggested times list fails. "); });

        }

    }

        // "9"
        // For switching from today to tomorrow
        // Note: Shortcut key need to be changed later
    else if (e.keyCode === 57) {

        // REPLACE UIA.ROOT()

        // Locate the Calendar Control on the top right
        return Q.fcall(function () { return uia.root().findFirst(function (el) { return (el.name == "Calendar Control"); }, 0, 11); })

        .then(function (calendarControlWindow) {

            calendarControlWindow.setFocus();

            // Find the date that is currently focused on Calendar Control
            /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
            // return calendarControlWindow.lastChild(); 
            /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////

            ////////////////////////HARDCODED
            return calendarControlWindow.firstChild().nextSibling().nextSibling();

        }, function (error) { throw new Error("The promise for finding Calendar Control window fails."); })
        .then(function (curDate) {
            var curDateStr = curDate.name;

            // Parse out the date number
            //////////////////////////////////////////////////////// ERROR HANDLING ////////////////////////////////////////////////////////
            cur_date_num = ((curDateStr.split(",")[1]).split(" "))[2];

            // Locate the dates table
            /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
            // return curDate.lastChild();
            /////////////////////////////////////////////////////////// BUGGY ///////////////////////////////////////////////////////////
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

            while (first_cell != null) {

                /////////////////////////////////////////////// ERROR CHECKING /////////////////////////////////////////////// 
                if (first_cell.name[1] === '1') {

                    break;
                }

                // Go through all the cells
                first_cell = first_cell.nextSibling();

            }

            // After finding the first cell of the month, we want to find the current day which has the name as curDateStr
            // NOTE: cur_cell must be a WEEKDAY. 

            // Keep finding
            while (first_cell.name != cur_date_num) {

                if (first_cell.nextSibling() != null) {

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
            if (first_cell.nextSibling() == null || first_cell.previousSibling() == null) { // it is weekend

                //////////////////////////////////////////////////// BREAK? How to? ////////////////////////////////////////////////////

                narrator.say("You've selected a weekend. Please select a weekday.");

            } else {
                var nextWeekday = find_next_sibling(first_cell);

                var nextWeekdayInvokePattern = nextWeekday.getPattern(10000);

                nextWeekdayInvokePattern.invoke(); // Click on it

                /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
                // Current implementation doesn't keep track of current day. This is different from the spec. 
                /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
            }

        }, function (error) { throw new Error("The promise for finding cells for the first week fails."); });

    }
   
    // Switching from today to yesterday
    // Key undefined
    else if (e.keyCode === 0) {


        ////////////////////////////////////////////////////// CHANGE /////////////////////////////////////////////////
        ///////////////////////////////////////////////// REFACTOR THE STEPS OF FINDING CURRENT CELL INTO A FUNCTION

        return Q.fcall(function () { return uia.root().findFirst(function (el) { return (el.name == "Calendar Control"); }, 0, 10) })

            .then(function (calendarControlWindow) {

            // Find the date that is currently focused on Calendar Control
            return calendarControlWindow.lastChild();


            }, function (error) { throw new Error("There is no Calandar Control!"); })

            .then(function (curDate) {
                var curDateStr = curDate.name;

                // Parse out the date number
                cur_date_num = ((curDateStr.split(",")[1]).split(" "))[1];

                // Locate the dates table
                return curDate.lastChild();

            }, function (error) { throw new Error("Can't find current date!"); })
            .then(function (datesTable) {

                // First week
                return datesTable.firstChild().nextSibling();

            }, function (error) { throw new Error("Can't find the dates table!"); })
            .then(function (firstWeekCellGroup) {

                // Locate the first one
                // Naive loop
                var first_cell = firstWeekCellGroup.firstChild();

                while (first_cell != null) {

                    if (first_cell.name == "1") {

                        break;
                    }

                    first_cell = first_cell.nextSibling(); // Go through to find the first day of the month

                }

                // After finding the first cell of the month, we want to find the current day which has the name as curDateStr
                // NOTE: cur_cell must be a WEEKDAY. 

                // Keep finding
                while (first_cell.name != cur_date_num) {

                    if (first_cell.nextSibling() != null) {

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
                if (first_cell.nextSibling() == null || first_cell.previousSibling() == null) { // it is weekend

                 //////////////////////////////////////////////////// BREAK? How to? ////////////////////////////////////////////////////

                    narrator.say("You've selected a weekend. Please select a weekday.");
                    return; // Can I just return in here?

                } else {
                    var nextWeekday = find_prev_sibling(first_cell);

                    var nextWeekdayInvokePattern = nextWeekday.getPattern(10000);

                    nextWeekdayInvokePattern.invoke(); // Click on it

                    /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
                    // Current implementation doesn't keep track of current day. This is different from the spec. 
                    /////////////////////////////////////////////////// NOTE ///////////////////////////////////////////////////
                }

            }, function (error) { throw new Error("Can't find the cells for the first week!"); });
    }
};
