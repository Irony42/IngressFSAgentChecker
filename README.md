# Ingress FS Agent checker

This script will help you to organize your event.

It will check : 
- Registered agents with no begin stats
- Agent with begin stats that are not registered on fevgames

- Agents that have begin stats but no end stats


## Installation
- Go to your IngressFS autoscore sheet.
- Open Tools -> Script editor
- Past the content of agentchecker.js (on this repository) on the editor and then save it.
- Go back to your sheet, on the "Data" tab.
- At the top you should have a new menu called "Agent checker". Open it and then open Check begin stats.
- The first time it might ask some authorizations to access the sheet data
- When prompted, past the list of the agents that registered on fevgames, separated bu spaces (for example : Agent01 Agent02 Agent03)


## Other
This script will not interfer with fevgame website. It don't send data anywhere.