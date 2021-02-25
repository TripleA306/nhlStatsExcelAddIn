(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This NHL stats addin is made by Aaron Atkinson and Jon Ronn");
                $('#button-text').text("Get Stats");
                $('#button-desc').text("Get your Stats");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }
            $('#content-main').css("background-image", "url('Images/hockeyRink.jpg')");
            $('#content-main').css("background-repeat", "no-repeat");
            $("#template-description").text("This NHL stats addin is made by Aaron Atkinson and Jon Ronn");
            $('#button-text').text("Show Team Stats");
            $('#button-desc').text("This button will show you the Team stats of the currently selected team.");

            $('#btn-text').text("Change Teams");
            $('#btn-desc').text("This button allows the user to change teams");

            //Will fill in the select box with NHL Team names
            populateNHLTeamsBox();
            $('#highlight-button').click(getTeamStatsAndRoster);
            $('#delete-sheet-button').click(clearSheets);

            //Hides the "Change Team" button as its not needed untill a team is selected
            $('#delete-sheet-button').hide();
        });
    };


    //This function will clear all of the sheets on the excel workbook (except the first sheet) and reload the web add-in
    //This is used to "fix" a issue with the event handlers.
    //If two NHL teams sheets were created, NONE of the table click handlers would work until a reload and a new team sheet was generated
    function clearSheets() {
        Excel.run(function (ctx) {
            let allSheets = ctx.workbook.worksheets;
            allSheets.load("items/name");

            //show the team select box
            $('#teamSelect').show("fast");

            return ctx.sync()
                .then(function () {
                    let numSheets = allSheets.items.length;
                    while (numSheets > 1) {

                        let lastSheet = allSheets.getLast();
                        lastSheet.delete();
                        numSheets--;
                    }

                    //reload the application when done
                    window.location.reload();
                    $('#highlight-button').show("fast");
                    return ctx.sync();
                });
        }).catch(errorHandler);
    }

    //This function will create a new sheet and display the selected NHL Team's Information and its current roster
    function displayTeamStats(teamInfo, roster) {

        var sheetName = teamInfo.name;

        //Show the "Change teams button
        $('#delete-sheet-button').show();
        //Set the h1 header to the teamName
        $('#teamName').text(sheetName);

        //Hide the NHL image and replace it with the NHL team Logo in 
        //hopes to hide the image from changing in a jarring way (3.5 seconds seemed to fix that)
        $('#nhlLogo').hide();
        $('#nhlLogo').attr('src', nhlLogoJSON.teams[0][teamInfo.id].url);
        $('#nhlLogo').fadeIn(3500);


        Excel.run(function (ctx) {

            //hide the team select box and Select Team Button
            $('#teamSelect').hide();
            $('#highlight-button').hide();

            let sheet = ctx.workbook.worksheets.add(sheetName);
            let teamRange = sheet.getRange("a1:b40").load("values, cellCount, format");
            let teamName = sheet.getRange("a1:b1").load("values, format");
            let statRange = sheet.getRange("a12:b41").load("values, format");

            let rosterTitle = sheet.getRange("f2:i2").load("values, format");
            let rosterDesc = sheet.getRange("f3:i3").load("values, format");

            //Merge the TeamName Cells
            teamName.merge();

            return ctx.sync()
                .then(function () {
                    //Grab the teamColor from our custom made JSON file 
                    let teamColor = nhlLogoJSON.teams[0][teamInfo.id].teamColor;

                    //Check to see if the contrast between black font and the team color is too low
                    // (not perfect but a decent attempt)
                    let tooDark = teamColor.substring(1, 3) <= 15 ? true : false;

                    //A counter designated for the row when populated team information
                    let counter = 2;

                    //Tried to make a single function when appling cell values to prevent duplicate code
                    //It would lose context inside the function and would overright data in the wrong column
                    for (var property in teamInfo) {

                        if (property === "name") {

                            let teamNameCell = teamName.getCell(0, 0);
                            teamNameCell.values = toTitleCase(sheetName);

                            teamName.format.font.bold = true;
                            teamName.format.font.size = 14;

                            teamName.format.horizontalAlignment = "Center";
                            teamName.format.fill.color = teamColor;

                            if (tooDark) {
                                teamNameCell.format.font.color = "white";
                            }

                        }

                        else if (property === "abbreviation" || property === "locationName"
                            || property === "firstYearOfPlay" || property === "officialSiteUrl") {

                            let descCell = teamRange.getCell(counter, 0);
                            descCell.format.font.bold = true;
                            let valCell = teamRange.getCell(counter, 1);
                            valCell.format.horizontalAlignment = "Left";
                            descCell.values = toTitleCase(property);
                            valCell.values = teamInfo[property];

                            counter++;
                        }
                        else if (property === "venue") {

                            let descCell = teamRange.getCell(counter, 0);
                            descCell.format.font.bold = true;
                            let valCell = teamRange.getCell(counter, 1);
                            valCell.format.horizontalAlignment = "Left";

                            descCell.values = "Arena";
                            valCell.values = teamInfo[property].name;
                            counter++;
                        }
                        else if (property === "division" || property === "conference") {

                            let descCell = teamRange.getCell(counter, 0);
                            descCell.format.font.bold = true;
                            let valCell = teamRange.getCell(counter, 1);
                            valCell.format.horizontalAlignment = "Left";

                            descCell.values = toTitleCase(property);
                            valCell.values = teamInfo[property].name;
                            counter++;
                        }
                        else if (property === "teamStats") {

                            //counter designated for the row when populated team stats information
                            let statCounter = 1;

                            //grabs ranged used for the Current Season stats Title and merges cells
                            let statTitleRange = sheet.getRange("A12:B12");
                            statTitleRange.merge();


                            let statTitleCell = statRange.getCell(0, 0);
                            statTitleCell.values = "Current Season Stats";

                            statTitleCell.format.font.bold = true;
                            statTitleCell.format.font.size = 14;
                            statTitleCell.format.horizontalAlignment = "Center";
                            statTitleCell.format.fill.color = teamColor;

                            if (tooDark) {
                                statTitleCell.format.font.color = "white";
                            }

                            //Get the stats object
                            let stats = teamInfo[property][0].splits[0].stat;

                            for (let stat in stats) {
                                let descCell = statRange.getCell(statCounter, 0);
                                descCell.format.font.bold = true;
                                let valCell = statRange.getCell(statCounter, 1);
                                valCell.format.horizontalAlignment = "Left";

                                descCell.values = toTitleCase(stat);
                                valCell.values = stats[stat];
                                statCounter++;
                            }
                        }
                    }

                    //Autofit the team information columns
                    teamRange.getEntireColumn().format.autofitColumns();

                    //Now create the Roster Table

                    let rosterTable = sheet.tables.add("E4:I4", true);
                    let tableName = "teamRoster" + teamInfo["abbreviation"];
                    rosterTable.name = tableName;
                    rosterTable.getHeaderRowRange().values = [["id", "Last Name", "First Name", "Jersey Number", "Position"]];

                    //Add the roster information to the table
                    for (let i = 0; i < roster.length; i++) {
                        let player = roster[i];
                        let id = player.person.id;
                        let fullname = player.person.fullName.split(" ");
                        let first = fullname[0];
                        let last = fullname[1];
                        let jersey = player.jerseyNumber;
                        let position = player.position.name;

                        rosterTable.rows.add(null, [[id, last, first, jersey, position]]);
                    }

                    //Sort the table by lastname
                    let sortRoster = rosterTable.getDataBodyRange();
                    sortRoster.sort.apply([
                        {
                            key: 1,
                            ascending: true,
                        },
                    ]);

                    //Format the table to have centered text for jesrsey number and make the header row team colors
                    rosterTable.columns.getItemAt(3).getDataBodyRange().format.horizontalAlignment = "Center";
                    rosterTable.getHeaderRowRange().format.fill.color = teamColor;

                    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
                        rosterTable.getRange().getEntireColumn().format.autofitColumns();
                    }

                    //Add the event handler to the table to allow for Player stat generation
                    rosterTable.onSelectionChanged.add(getAndDisplayPlayerStats);

                    //Hide the ID range
                    sheet.getRange("e4").getEntireColumn().columnHidden = true;

                    //Add and format Roster title and sub header to notify the user of its functionality
                    rosterTitle.getCell(0, 0).values = [["Roster"]];
                    rosterDesc.getCell(0, 0).values = [["( Click on a player to view their stats )"]]
                    rosterTitle.getCell(0, 0).format.font.bold = true;
                    rosterTitle.format.font.size = 14;
                    rosterTitle.format.horizontalAlignment = "Center";
                    rosterTitle.format.fill.color = teamColor;

                    if (tooDark) {
                        rosterTitle.format.font.color = "white";
                    }

                    rosterDesc.format.horizontalAlignment = "Center";
                    rosterDesc.merge();
                    rosterTitle.merge();

                    //Change the active sheet to the NHL team sheet
                    sheet.activate();


                }).then(ctx.sync);

        }).catch(errorHandler);
    }


    //This function is used to convert the camel case information from the api to a more User Friendly format
    function toTitleCase(string) {

        if (isNaN(string)) {
            if (string.length > 2) {
                string = string.replace(/\w/, c => c.toUpperCase());
                string = string.replace(/([a-z])([A-Z])/g, '$1 $2');
                string = string.replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
            }
            else if (string === "ot") { //The word is a two letter acronym
                string = string.toUpperCase();
            }
        }

        return string;
    }

    //This function gets called when a user clicks on the Roster Table
    //It will grab information from the API and display the Players Information
    //And also their Year by Year stats on a new sheet
    function getAndDisplayPlayerStats(event) {

        let selection = event.address.split(":");
        //get the player row number to identify which player was selected
        let playerRow = selection[0];
        let insideTable = event.isInsideTable;
        //Used to grab the table in excel.run(), as we lose context in the click handler
        let tableID = event.tableId;

        if (insideTable) {

            Excel.run(function (ctx) {
                let sheet = ctx.workbook.worksheets.getActiveWorksheet();
                let rosterTable = sheet.tables.getItem(tableID);
                let idColumn = rosterTable.columns.getItem(1).load("values");

                //Create a new Sheet with a temporary name thats slightly unique, indicated to the user that the information is being processed 
                let newSheet = ctx.workbook.worksheets.add("Loading....Player" + playerRow.substring(1, playerRow.length + 1));

                //Grabbed an exaggerated range, some people have alot more information than others.
                let playerStatsRange = newSheet.getRange("A1:ZZ25").load("values, format");

                /**
                    !!!THIS IS OUR CROWNING ACHIEVEMENT!!!!
                    Due to the nature of the event handler there was no way for us to load player information into a secondary page without
                    loading all of the player information for each person before displaying the teamSheet (not a good UX idea)

                    Due to the nature of this function, the context of the playerStatsRange will be lost when ctx.sync() is run.
                    ctx.trackedObjects.add() allows you to store an excel object into memory to be used outside of its original context.
                    
                    Once changes are made to the excel object outside of its range <excelObjName>.context.sync(); must be called to update the values.
                    And just like C class.. we must free the memory when we are finished with it, as this data is stored throughout the workbooks life span.

                    We will take our 100% now please :)
                */
                ctx.trackedObjects.add(playerStatsRange);

                return ctx.sync()
                    .then(function () {

                        //grab the playerID from the Roster Table (hidden column)
                        let playerID = idColumn.values[parseInt(playerRow.substring(1, playerRow.length + 1)) - 4][0];


                        $.when(getPlayerInfo(playerID),
                            getPlayerStats(playerID))
                            .done(function (playerData, statData) {

                                let playerInfo = playerData[0].people[0];
                                let playerStats = statData[0].stats[0].splits;

                                //Change the name of the sheet to the Players name
                                newSheet.name = playerInfo.fullName;

                                let counter = 0;

                                //Remove the attributes we dont want
                                delete playerInfo['id'];
                                delete playerInfo['firstName'];
                                delete playerInfo['lastName'];
                                delete playerInfo['link'];
                                delete playerInfo['currentTeam'];
                                delete playerInfo['primaryPosition'];


                                //Loop through and display all of the player information
                                for (let stat in playerInfo) {

                                    let descCell = playerStatsRange.getCell(counter, 0);
                                    descCell.format.font.bold = true;
                                    let valCell = playerStatsRange.getCell(counter, 1);
                                    valCell.format.horizontalAlignment = "Left";

                                    if (stat === "active" || stat === "alternateCaptain" || stat === "captain" || stat === "rookie") {
                                        playerInfo[stat] = playerInfo[stat] === true ? "Yes" : "No";
                                    }

                                    descCell.values = toTitleCase(stat);
                                    valCell.values = toTitleCase(playerInfo[stat]);
                                    counter++;
                                }

                                //used to append the starting stat column to the column D
                                let statColumn = 3

                                //Loop through and display all of the Players stats, placed vertically by season
                                for (let i = 0; i < playerStats.length; i++) {

                                    let statCounter = 1;
                                    //format the year from yyyy-yyyy to yyyy - yyyy
                                    let year = playerStats[i].season.slice(0, 4) + " - " + playerStats[i].season.slice(4);

                                    let yearCell = playerStatsRange.getCell(0, statColumn);
                                    yearCell.format.font.bold = true;
                                    let numCell = playerStatsRange.getCell(0, statColumn + 1);
                                    numCell.format.horizontalAlignment = "Left";
                                    numCell.format.font.bold = true;

                                    yearCell.values = "Season";
                                    numCell.values = year;

                                    let stats = playerStats[i].stat;

                                    for (let stat in stats) {
                                        let descCell = playerStatsRange.getCell(statCounter, statColumn);
                                        let valCell = playerStatsRange.getCell(statCounter, statColumn + 1);
                                        valCell.format.horizontalAlignment = "Left";

                                        descCell.values = toTitleCase(stat);
                                        valCell.values = stats[stat];
                                        statCounter++;
                                    }

                                    //append next season to the previous season with a nice space in-between
                                    statColumn += 3;

                                }

                                //AutoFit the Columns now that we're done
                                playerStatsRange.getEntireColumn().format.autofitColumns();

                                //Sync up the changes made to the player Range, and free the memory
                                playerStatsRange.context.sync();
                                playerStatsRange.context.trackedObjects.remove();
                            });
                    }).then(ctx.sync);

            }).catch(errorHandler)

        }
    }

    //This function will populate the select box will all of the NHL teams 
    function populateNHLTeamsBox() {
        let Teams = [];

        getTeams().done(function (data) {

            //Sort the teams by team name (was id)
            data.teams.sort(function (a, b) {
                if (a.name < b.name) {
                    return -1;
                }
                if (a.name > b.name) {
                    return 1;
                }
                return 0
            });

            Teams = data.teams;

            //Sort team names

            for (let i = 0; i < Teams.length; i++) {

                //Add the team to the select box
                $('#teamSelect').append(new Option(Teams[i].name + "", Teams[i].id + ""));

            }
        });
    }

    //This function will get the JSON from the team and Roster API call and run the displayTeamStats function
    function getTeamStatsAndRoster() {

        //get the value from the corresponding team in the select box
        let selectedID = $('#teamSelect').val();

        if (selectedID > -1) {
            let teamUrl = "https://statsapi.web.nhl.com/api/v1/teams/" + selectedID + "?expand=team.stats";
            let rosterUrl = "https://statsapi.web.nhl.com/api/v1/teams/" + selectedID + "?expand=team.roster";

            $.when(
                $.getJSON(teamUrl),
                $.getJSON(rosterUrl)
            ).done(function (teamData, rosterData) {
                displayTeamStats(teamData[0].teams[0], rosterData[0].teams[0].roster.roster);
            })

        }
        else {
            showNotification("Please select a Team!");
        }
    }

    //Returns the promise of the Teams GET request
    function getTeams() {
        let url = "https://statsapi.web.nhl.com/api/v1/teams?expand=person.names";
        return $.getJSON(url);
    }

    //Returns the promise of the Player Info GET request
    function getPlayerInfo(id) {

        //In case the user clicks on the headers
        if (id !== "id") {
            let url = "https://statsapi.web.nhl.com/api/v1/people/" + id;

            return $.getJSON(url);
        }
    }

    //Returns the promise of the Player All-Time Stats GET request
    function getPlayerStats(id) {
        //In case the user clicks on the headers
        if (id !== "id") {
            let url = "https://statsapi.web.nhl.com/api/v1/people/" + id + "/stats/?stats=yearByYear";

            return $.getJSON(url);
        }
    }


    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    //a JSON created by us to grab Logo URL's and TeamColor based on the API's team ID
    //Grabbing the JSON from a local file caused issues
    var nhlLogoJSON =
    {
        teams: [
            {
                "1": {
                    "team": "new jersey",
                    "url": "https://assets.detroithockey.net/dyn/4d/0f9e3c5d1d402290cc4e83e50bcb67.png",
                    "teamColor": "#CE1126"
                },
                "2": {
                    "team": "new york islanders",
                    "url": "https://assets.detroithockey.net/dyn/d6/c41f42e2252a88d1d8bbca580cb6e7.png",
                    "teamColor": "#00539B"
                },
                "3": {
                    "team": "new york rangers",
                    "url": "https://assets.detroithockey.net/dyn/6e/c094260e542eab65eb921fdb33f481.png",
                    "teamColor": "#0038A8"
                },
                "4": {
                    "team": "phili",
                    "url": "https://assets.detroithockey.net/dyn/33/2f5d68205136ef97caab7ef9fdc9ae.png",
                    "teamColor": "#F74902"
                },
                "5": {
                    "team": "pittsburg",
                    "url": "https://assets.detroithockey.net/dyn/3d/c15f61b020bb80ed80ed288e6586f6.png",
                    "teamColor": "#FCB514"
                },
                "6": {
                    "team": "boston",
                    "url": "https://assets.detroithockey.net/dyn/80/069389c92dc4e9048952358e0bcc70.png",
                    "teamColor": "#FFB81C"
                },
                "7": {
                    "team": "buffalo",
                    "url": "https://assets.detroithockey.net/dyn/01/fb4b8625323ae8e825887806adf8a3.png",
                    "teamColor": "#002654"
                },
                "8": {
                    "team": "montreal",
                    "url": "https://assets.detroithockey.net/dyn/02/ce59fd44a46d4d84e10cd5e2bf72a6.png",
                    "teamColor": "#AF1E2D"
                },
                "9": {
                    "team": "ottawa",
                    "url": "https://assets.detroithockey.net/dyn/91/08b7060146b629c4d91a5f65a1876b.png",
                    "teamColor": "#C52032"
                },
                "10": {
                    "team": "toronto",
                    "url": "https://assets.detroithockey.net/dyn/91/77814246cca5cd69f524f82dddfea4.png",
                    "teamColor": "#00205B"
                },
                "12": {
                    "team": "carolina",
                    "url": "https://assets.detroithockey.net/dyn/e7/7fcaa1e765302100b1513f1daf80fa.png",
                    "teamColor": "#CC0000"
                },
                "13": {
                    "team": "florida",
                    "url": "https://assets.detroithockey.net/dyn/ce/32a8af9a69a2b96c58bd480229897a.png",
                    "teamColor": "#041E42"
                },
                "14": {
                    "team": "tbay",
                    "url": "https://assets.detroithockey.net/dyn/5b/9ed793293913292a2a9e948ee49a20.png",
                    "teamColor": "#002868"
                },
                "15": {
                    "team": "washington",
                    "url": "https://assets.detroithockey.net/dyn/16/c9c491eb59c2874e5fb7084c544538.png",
                    "teamColor": "#C8102E"
                },
                "16": {
                    "team": "chicago",
                    "url": "https://assets.detroithockey.net/dyn/e7/7fcaa1e765302100b1513f1daf80fa.png",
                    "teamColor": "#CF0A2C"
                },
                "17": {
                    "team": "detroit",
                    "url": "https://assets.detroithockey.net/dyn/a6/1b1dcdce6fe656ec83d296876718ee.png",
                    "teamColor": "#CE1126"
                },
                "18": {
                    "team": "nashville",
                    "url": "https://assets.detroithockey.net/dyn/6c/6b77f8ed9065a3672f5d53b0df6393.png",
                    "teamColor": "#FFB81C"
                },
                "19": {
                    "team": "st louis",
                    "url": "https://assets.detroithockey.net/dyn/2f/92d32b1524d577eddde18f56470564.png",
                    "teamColor": "#002F87"
                },
                "20": {
                    "team": "calgary",
                    "url": "https://assets.detroithockey.net/dyn/49/643ffd7dae95e9ef34c877f991a86e.png",
                    "teamColor": "#C8102E"
                },
                "21": {
                    "team": "colorado",
                    "url": "https://assets.detroithockey.net/dyn/7d/f4f33a640140d1b668bd6d44f6dd66.png",
                    "teamColor": "#6F263D"
                },
                "22": {
                    "team": "edmonton",
                    "url": "https://assets.detroithockey.net/dyn/bf/47567b767c0a920c613fbbb2a2c4bf.png",
                    "teamColor": "#041E42"
                },
                "23": {
                    "team": "vancouver",
                    "url": "https://assets.detroithockey.net/dyn/f6/b302a063b5b3497e3a2a584ec85f50.png",
                    "teamColor": "#0038A8"
                },
                "24": {
                    "team": "anaheim",
                    "url": "https://assets.detroithockey.net/dyn/93/557f13cb4ebcf23a82e205abed79c7.png",
                    "teamColor": "#F47A38"
                },
                "25": {
                    "team": "dallas",
                    "url": "https://assets.detroithockey.net/dyn/0e/def4ee389b4504764eb6f3cc8648b4.png",
                    "teamColor": "#006847"
                },
                "26": {
                    "team": "la",
                    "url": "https://assets.detroithockey.net/dyn/b5/3a131b974648ddcac21b0a12ce9a5c.png",
                    "teamColor": "#572A84"
                },
                "28": {
                    "team": "san jose",
                    "url": "https://assets.detroithockey.net/dyn/50/5d2b7ac504100a9a66b7215c0d1552.png",
                    "teamColor": "#006D75"
                },
                "29": {
                    "team": "columbus",
                    "url": "https://assets.detroithockey.net/dyn/85/e99801f8b5526f31f312de1393bafc.png",
                    "teamColor": "#002654"
                },
                "30": {
                    "team": "minnesota",
                    "url": "https://assets.detroithockey.net/dyn/a0/3ac464e4162de246a12f064249712c.png",
                    "teamColor": "#154734"
                },
                "52": {
                    "team": "winnipeg",
                    "url": "https://assets.detroithockey.net/dyn/c6/f50bd92f4ef8be45b03b9a44c94a4f.png",
                    "teamColor": "#041E42"
                },
                "53": {
                    "team": "arizona",
                    "url": "https://assets.detroithockey.net/dyn/7b/f26f775df6df90510823977ebe1219.png",
                    "teamColor": "#8C2633"
                },
                "54": {
                    "team": "vegas",
                    "url": "https://assets.detroithockey.net/dyn/32/3a443b578e6873bff5a36805141cdd.png",
                    "teamColor": "#B4975A"
                }
            }
        ]
    }

})();


