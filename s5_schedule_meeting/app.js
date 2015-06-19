/// <reference path="uia.d.ts" />

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


function continueRoomSelection(desktop) {

    return Q.fcall(function () {
        return desktop.findFirst(function (el) { return (el.name.indexOf("Select Rooms") > -1); }, 0, 2);
    }).then(function (selRoomWindow) {

        // Check More columns
        return Q.fcall(function () {
            return selRoomWindow.findFirst(function (el) { return (el.name == "More columns"); }, 0, 1);
        }).then(function (moreColumnsButton) {

            // Click on the button
            var invokePattern = moreColumnsButton.getPattern(10000);
            invokePattern.invoke();

            // Question: Do we have to wait?
            return selRoomWindow.findFirst(function (el) { return ((el.name == "Search:") && (el.className == "RichEdit20W")); }, 0, 1);

        }, function (error) { throw new Error("Can't locate More columns radio button!"); })
        .then(function (editBox) {

            // Begin editing
            ////////////////////////////////////////// EATEN //////////////////////////////////////////
            // narrator.say("Type the building number or any keyword about the room. After you finish, press enter to start searching.");

            // Set the focus on edit box
            editBox.setFocus();

            //////////////////// After this, add a shortcut key for user to press the OK button directly ////////////////////

        }, function (error) { throw new Error("Can't find edit box!"); });

    }, function (error) { throw new Error("Can't locate the window for selecting rooms!"); });

}

//////////////////////////////////////////// Global variable ////////////////////////////////////////////
var cur_suggestion_item; // Global var for reading out suggestion times

host.onKeypress = function (e) {
    console.log("onkeypress");
    console.log(JSON.stringify(e));

    // "1"
    if (e.keyCode === 49) {
        //narrator.say("1 2 3");

    }

    // "2"
    else if (e.keyCode === 50) {

    }

        // "4"
        // Scenario 4
    else if (e.keyCode === 52) {
        ///////////////////////////////////// Code for scenario 4 /////////////////////////////////////
    }

    // "5"
    else if (e.keyCode === 53) {
        debugger;
    }

        // "6"
        // Scenario for scheduling a meeting 
        // Problem: Don't know how to click on the Calendar button though
    else if (e.keyCode === 54) {

        return Q.fcall(function () {

            // Find desktop element
            return uia.root();

        }).then(function (desktop) {

            // Suppose you are already in the window of scheduling a new meeting, find
            // "Scheduling Assistant" button
            return Q.fcall(function () {
                return desktop.findFirst(function (el) { return (el.name == "Scheduling Assistant"); }, 0, 11)
            }).then(function (schedulingAssistantButton) {

                // Click on the button using toggle pattern... maybe?
                var togglePat = schedulingAssistantButton.getPattern(10015);
                togglePat.toggle();

                // inform the user
                // Maybe after waiting for 2 seconds?
                /////////////////////////////////////////////// EATEN ///////////////////////////////////////////////
                narrator.say("You are in the scheduling assistant tab. Please choose attendees and room. ");
                /////////////////////////////////////////////// EATEN ///////////////////////////////////////////////


            }, function (error) { throw new Error("Can't find Scheduling Assistant tab item! "); });

        }, function (error) { throw new Error("Can't find desktop element! "); });

    }

        // After getting into Scheduling Assistant tab, we can do the following stuffs:
        // Press "7"
        // The shortcut key for adding rooms
    else if (e.keyCode === 55) {

        return Q.fcall(function () {
            return uia.root();
        }).then(function (desktop) {

            // Locate "Add rooms..." button
            return Q.fcall(function () {

                return desktop.findFirst(function (el) { return (el.name.indexOf("Add Rooms") > -1); }, 0, 5);

            }).then(function (addRoomsButton) {

                console.log(addRoomsButton.name);

                console.log(addRoomsButton.getProperty(30031));

                //// Get Invoke pattern
                ////////////////////////////////////////// BREAKS ////////////////////////////////////////
                //var invokePattern = addRoomsButton.getPattern(10000);
                //// Press it
                //invokePattern.invoke();
                ////////////////////////////////////////// BREAKS ////////////////////////////////////////

                // Wait for 2 seconds and continue
                host.setTimeout(function () { continueRoomSelection(desktop); }, 2000);

            }, function (error) {
                narrator.say("Check if you already opened Scheduling Assistant.");
                throw new Error("Can't find Add Rooms... button!");
            });

        }, function (error) { throw new Error("Can't find desktop button!"); });

    }
        // "8"
        // After the user selects room and attendees, for reading out suggested times
    else if (e.keyCode === 56) {

        return Q.fcall(function () {
            return uia.root();
        }).then(function (desktop) {

            // Locate suggested time list
            return desktop.findFirst(function (el) { return ((el.name == "Suggested times:") && (el.className == "ATL:00007FF92289D640")); }, 0, 10);

        }, function (error) { throw new Error("Can't find desktop element!"); })
        .then(function (suggestedTimesList) {

            // In the list, we first locate on the first item
            // At this time, the Narrator might read out something
            // Warning: So, if we now let Narrator say something, it might be eaten

            /////////////////////////////////////// CHANGE LATER ///////////////////////////////////////
            cur_suggestion_item = suggestedTimesList.firstChild().nextSibling(); // FIRST CHILD.
            /////////////////////////////////////// CHANGE LATER ///////////////////////////////////////

            // Reading out the first child
            if (cur_suggestion_item != null) {

                console.log("First room: " + cur_suggestion_item.name);

                // Reads out time and name
                narrator.say(cur_suggestion_item.name);

                cur_suggestion_item.setFocus(); // Set focus to it

                if (cur_suggestion_item.name.indexOf("No available rooms") <= -1) {

                    // There is available rooms
                    // Check if there is a conflict
                    if (cur_suggestion_item.name.indexOf("conflict") > -1) {
                        // If there is a conflict
                        // Then we read it out -- Only the attendees who have a conflict
                        // Find the window

                        return Q.fcall(function () {
                            return uia.root().findFirst(function (el) {
                                return ((el.name.indexOf("All Attendees Status") > -1) && (el.className == "AfxWndW"));
                            }, 0, 5);
                        }).then(function (attendeeStatusWindow) {

                            console.log("get here");

                            var attendeeStatus = attendeeStatusWindow.name;

                            var statusArray = attendeeStatus.split(';');

                            // Reading out the status --> Including rooms
                            for (var i = 2; i < statusArray.length; i++) {

                               
                                if (statusArray[i].indexOf("Free") <= -1) { // Not free = conflict

                                    //////////////////////////////////////////////// PROBLEM ////////////////////////////////////////////////
                                    console.log("Array member: " + statusArray[i]);
                                    narrator.stopReading(); // NO USE
                                    narrator.say(statusArray[i]); // EATEN
                                    //////////////////////////////////////////////// PROBLEM ////////////////////////////////////////////////
                                }
                            }

                        }, function (error) { throw new Error("Can't find attendees status window!"); })

                    }

                }


            } else {
                narrator.say("There is no suggested times.");
            }

        }, function (error) { throw new Error("Can't find suggested times list!"); })

    }

    //////////////////////////////////////////////// TBA: NEED KEYCODE ////////////////////////////////////////////////
//        // "d"
//        // The shortcut keys for going through the children of suggested times list --> GOING DOWN
//    else if (e.keyCode === ) {


//    }

//    // "u"
//    // The shortcut keys for going through the children of suggested times list --> GOING UP
//else if (e.keyCode === ) {

//}
};
