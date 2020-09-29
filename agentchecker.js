var ss = SpreadsheetApp.getActiveSpreadsheet();

function onInstall() {
    onOpen();
}

function onOpen() {
    ss.addMenu("Agent checker", [
        { name: "Check begin stats", functionName: "beginStats" },
        { name: "Check end stats", functionName: "endStats" }]);
}

function beginStats() {
    var sheet = ss.getActiveSheet();
    var agentsWithoutStats = [];
    var nonRegisteredAgents = [];
    var agentsWithBeginStats = sheet.getRange("E:E").getDisplayValues();
    var registeredAgents = [];

    var names = Browser.inputBox("Agent list", "Fill with the agent list, separated with spaces. For example : Agent01 Agent02 Agent03", Browser.Buttons.OK);
    registeredAgents = names.split(" ");

    registeredAgents.forEach(registeredAgent => {
        var index = agentsWithBeginStats.findIndex(agent => agent.length > 0 ? agent[0].toLowerCase() == registeredAgent.toLowerCase() : false);
        if (index == -1) {
            agentsWithoutStats.push(registeredAgent);
        }
    })
    agentsWithBeginStats.shift(); //remove column title
    agentsWithBeginStats.forEach(agentWithBeginStats => {
        if(agentWithBeginStats.length > 0) {
            var index = registeredAgents.findIndex(registeredAgent => registeredAgent.toLowerCase() == agentWithBeginStats[0].toLowerCase());
            if (index == -1 && typeof agentWithBeginStats[0] === "string" && agentWithBeginStats[0] !== "") {
                nonRegisteredAgents.push(agentWithBeginStats[0]);
            }
        }
    })
    Browser.msgBox("Registered agents with no begin stats: " + agentsWithoutStats, Browser.Buttons.OK);
    Browser.msgBox("Non registered agents with begin stats: " + nonRegisteredAgents, Browser.Buttons.OK);
}

function endStats() {
    var sheet = ss.getActiveSheet();
    var agentNameCol = sheet.getRange("E:E").getDisplayValues();
    var endLevelCol = sheet.getRange("I:I").getDisplayValues();
    var agentWithoutEndStats = [];
    for (var i = 1; i < agentNameCol.length; i++) {
        var agentName = agentNameCol[i];
        if (agentName && agentName.length > 0 && typeof agentName[0] === "string" && agentName[0] !== "") {
            var agentEndLevel = endLevelCol[i];
            if (!(agentEndLevel > 0 && agentEndLevel < 17)) {
                agentWithoutEndStats.push(agentName);
            }
        }
    }
    Browser.msgBox("Agents with some begin stats but no end stats: " + agentWithoutEndStats, Browser.Buttons.OK);
}
