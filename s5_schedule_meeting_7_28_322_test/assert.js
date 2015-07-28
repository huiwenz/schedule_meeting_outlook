
// Reference: A JS assertion function from http://stackoverflow.com/questions/15313418/javascript-assert
// Written by myself
function assertStrictEqual(obj1, obj2, testTitle) {

    // Suppose the message will only be displayed 
    // when the assertion is true
    if (obj1 !== obj2) {
        console.log((testTitle || "Assertion for strict equality") + " failed: You are comparing " + obj1 + " and " + obj2 + ".");
    } else {
        console.log((testTitle || "Assertion for strict equality") + " passed! ");
        
    }
}

// Check if the condition meets.
function assertOk(condition, testTitle) {

    if (!condition) {
        console.log((testTitle || "Assertion for a true statement") + " failed.");
    } else {
        console.log((testTitle || "Assertion for a true statement") + " passed! ");
    }

}
