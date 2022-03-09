const _ = LodashGS.load();

// tournament section header
const TOURNAMENT_SECTION_HEADERS = 
  [
    {
      displayName: 'Name',
      key: 'displayName'
    }, 
    {
      displayName: 'Link',
      key: 'tournamentLink'
    }, 
    {
      displayName: 'ID',
      key: 'tournamentId' 
    },
    {
      displayName: 'Games Number',
      formula: getGamesNumberFormula
    },
    {
      displayName: 'Start Date',
      formula: getStartDateFormula
    }, 
    {
      displayName: 'End Date',
      formula: getEndDateFormula
    },
    {
      displayName: 'Average Difference',
      formula: getAverageDifferenceFormula
    }
  ];

// tournament section games
const GAME_HEADERS = 
  [
    {
      displayName: 'TEAM A',
      key: 'teamACode'
    }, 
    {
      displayName: 'TEAM B',
      key: 'teamBCode'
    },
    {
      displayName: 'Game Date',
      key: 'gameDate'
    },
    {
      displayName: 'TEAM A Score',
      key: 'teamAScore'
    }, 
    {
      displayName: 'TEAM B Score',
      key: 'teamBScore'
    }, 
    {
      displayName: 'Winner',
      formula: getWinnerFormula
    },
    {
      displayName: 'Difference',
      formula: getDifferenceFormula
    }
];
//['TEAM A', 'TEAM B', 'Game Date', 'TEAM A Score', 'TEAM B Score', 'Winner', 'Difference'];

const TOURNAMENTS_SECTIONS_FIRST_ROW = 1;
const TOURNAMENTS_SECTIONS_FIRST_COLUMN = 1;
const TOURNAMENTS_SECTIONS_PADDING = 2;
const TOURNAMENTS_SECTIONS_STEP = GAME_HEADERS.length + TOURNAMENTS_SECTIONS_PADDING;

const GAMES_SECTIONS_TITLE = 'Games'
const GAMES_SECTIONS_PADDING = 1;
const GAMES_SECTIONS_OFFSET_FROM_START = TOURNAMENTS_SECTIONS_FIRST_ROW + TOURNAMENT_SECTION_HEADERS.length + GAMES_SECTIONS_PADDING + 1;
const GAMES_SECTIONS_OFFSET_FROM_SECTION = TOURNAMENT_SECTION_HEADERS.length + GAMES_SECTIONS_PADDING + 1;

// config
const CONFIG_TOURNAMENT_HEADERS = ['Selected', 'Tournament ID', 'Tournament Name', 'Tournament Link', 'Games Number', 'Start Date', 'End Date', 'Competition', 'Name To Show', 'Active', 'Order']
const CONFIG_TOURNAMENTS_SECTION_FIRST_ROW = 3;
const CONFIG_TOURNAMENTS_SECTION_FIRST_COLUMN = 1;

// main
const MAIN_SECTIONS_FIRST_ROW = 1;
const MAIN_SECTIONS_FIRST_COLUMN = 1;
const MAIN_TOURNAMENTS_SECTION_HEADERS = [
  {
    displayName: 'Name',
    tournamentHeaderDisplayName: 'Name'
  },
  {
    displayName: 'Games',
    tournamentHeaderDisplayName: 'Games Number'
  },
  {
    displayName: 'Average Difference',
    tournamentHeaderDisplayName: 'Average Difference'
  }
];
const MAIN_TOURNAMENTS_SECTIONS_PADDING = 1;
const MAIN_SECTIONS_PADDING = 2;
const MAIN_SECTIONS_STEP = MAIN_TOURNAMENTS_SECTION_HEADERS.length + MAIN_SECTIONS_PADDING;

const CONTROLLER_LAST_REFRESH_CELL = 'K14';
const CONTROLLER_LAST_RESET_CELL = 'L14';

function getStartDateFormula(sheet, sectionFirstRow, sectionFirstColumn){
  let firstGameDateRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
  let firstGameDateColumn = sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'Game Date');
  let datesRange = getColumnToEnd(sheet.getRange(firstGameDateRow, firstGameDateColumn, 2, 1).getA1Notation());
  return `=sortn(${datesRange},1,0,1,1)`;
}

function getEndDateFormula(sheet, sectionFirstRow, sectionFirstColumn){
  let firstGameDateRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
  let firstGameDateColumn = sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'Game Date');
  let datesRange = getColumnToEnd(sheet.getRange(firstGameDateRow, firstGameDateColumn, 2, 1).getA1Notation());
  return `=sortn(${datesRange},1,0,1,0)`;
}

function getAverageDifferenceFormula(sheet, sectionFirstRow, sectionFirstColumn){
  let firstGameDiffRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
  let firstGameDiffColumn = sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'Difference');
  let diffsRange = getColumnToEnd(sheet.getRange(firstGameDiffRow, firstGameDiffColumn, 2, 1).getA1Notation());
  return `=IFERROR(round(averageif(${diffsRange}, "<>0"), 2),"")`;
}

function getGamesNumberFormula(sheet, sectionFirstRow, sectionFirstColumn){
  let firstGameDateRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
  let firstGameDateColumn = sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'Game Date');
  let datesRange = getColumnToEnd(sheet.getRange(firstGameDateRow, firstGameDateColumn, 2, 1).getA1Notation());
  return `=counta(${datesRange})`;
}

function getWinnerFormula(sheet, sectionFirstRow, sectionFirstColumn, gameIndex){
  let gameRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1 + gameIndex;
  let teamAScoreCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM A Score'), 1, 1).getA1Notation();
  let teamBScoreCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM B Score'), 1, 1).getA1Notation();
  let teamACodeCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM A'), 1, 1).getA1Notation();
  let teamBCodeCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM B'), 1, 1).getA1Notation();
  return `=if(${teamAScoreCell} > ${teamBScoreCell}, ${teamACodeCell}, ${teamBCodeCell})`;
}

function getDifferenceFormula(sheet, sectionFirstRow, sectionFirstColumn, gameIndex){
  let gameRow = sectionFirstRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1 + gameIndex;
  let teamAScoreCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM A Score'), 1, 1).getA1Notation();
  let teamBScoreCell = sheet.getRange(gameRow, sectionFirstColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM B Score'), 1, 1).getA1Notation();
  return `=if(not(isblank(${teamAScoreCell})), abs(${teamAScoreCell} - ${teamBScoreCell}), "")`;
}

function getColumnToEnd(a1Notation){
  return a1Notation.replace(/\d+$/, "");
}

function getRowToEnd(a1Notation){
  return a1Notation.replace(/[A-Z]+(?=\d*$)/, "");
}

function ensureSheetExist(spreadsheet, sheetName){
  var itt = spreadsheet.getSheetByName(sheetName);
  if (!itt) {
    spreadsheet.insertSheet(sheetName);
  }
  return spreadsheet.getSheetByName(sheetName);
}

function getAllTournaments(){
  const jsonString = HtmlService.createHtmlOutputFromFile("tournaments.html").getContent();
  return JSON.parse(jsonString);
}

function getCurrentTournamentsConfig(allwaysActive=false){
  var sheetApp = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ensureSheetExist(sheetApp, 'Config');
  let tournamentsArrays = configSheet.getSheetValues(CONFIG_TOURNAMENTS_SECTION_FIRST_ROW + 1, CONFIG_TOURNAMENTS_SECTION_FIRST_COLUMN, -1, CONFIG_TOURNAMENT_HEADERS.length);
  let tournamentsDict = Object.assign({}, ...(tournamentsArrays.map(t => ({
    [t[CONFIG_TOURNAMENT_HEADERS.indexOf('Tournament ID')]]: {
      selected: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Selected')],
      tournamentId: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Tournament ID')],
      tournamentName: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Tournament Name')],
      tournamentLink: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Tournament Link')],
      gamesCount: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Games Number')],
      startDate: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Start Date')],
      endDate: t[CONFIG_TOURNAMENT_HEADERS.indexOf('End Date')],
      competition: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Competition')],
      displayName: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Name To Show')],
      active: allwaysActive || t[CONFIG_TOURNAMENT_HEADERS.indexOf('Active')],
      order: t[CONFIG_TOURNAMENT_HEADERS.indexOf('Order')]
    }
  }))));
  return tournamentsDict;
}

function refreshConfigSheet(){
  let tournamenets = getAllTournaments();
  let currentConfig = getCurrentTournamentsConfig();
  let tournamentsArrays = tournamenets.map((t, i) => 
  [
    currentConfig[t.tournamentId] ? currentConfig[t.tournamentId].selected : false,
    t.tournamentId, 
    t.tournamentName, 
    t.tournamentLink, 
    t.gamesCount, 
    new Date(t.startDateEpoch).toISOString(), 
    new Date(t.endDateEpoch).toISOString(), 
    currentConfig[t.tournamentId] ? currentConfig[t.tournamentId].competition : '', 
    currentConfig[t.tournamentId] ? currentConfig[t.tournamentId].displayName : t.tournamentName, 
    currentConfig[t.tournamentId] ? currentConfig[t.tournamentId].active : false, 
    currentConfig[t.tournamentId] ? currentConfig[t.tournamentId].order : i
  ]);
  setConfigData(tournamentsArrays);
}

function resetConfigSheet(){
  let tournamenets = getAllTournaments();
  let tournamentsArrays = tournamenets.map((t, i) => 
  [
    false,
    t.tournamentId, 
    t.tournamentName, 
    t.tournamentLink, 
    t.gamesCount, 
    new Date(t.startDateEpoch).toISOString(), 
    new Date(t.endDateEpoch).toISOString(), 
    '', 
    t.tournamentName, 
    false, 
    i
  ]);
  setConfigData(tournamentsArrays);
}

function setConfigData(tournamentsArrays){
  var sheetApp = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ensureSheetExist(sheetApp, 'Config');
  let toSave = [CONFIG_TOURNAMENT_HEADERS].concat(tournamentsArrays);
  let configRange = configSheet.getRange(CONFIG_TOURNAMENTS_SECTION_FIRST_ROW, CONFIG_TOURNAMENTS_SECTION_FIRST_COLUMN, toSave.length, CONFIG_TOURNAMENT_HEADERS.length);
  configSheet.getRange(1, 1, configSheet.getMaxRows(), configSheet.getMaxColumns()).clearDataValidations();
  configSheet.clear();
  configRange.setValues(toSave);
  let selectedRange = configSheet.getRange(CONFIG_TOURNAMENTS_SECTION_FIRST_ROW + 1, CONFIG_TOURNAMENT_HEADERS.indexOf('Selected') + 1, tournamentsArrays.length, 1);
  let activeRange = configSheet.getRange(CONFIG_TOURNAMENTS_SECTION_FIRST_ROW + 1, CONFIG_TOURNAMENT_HEADERS.indexOf('Active') + 1, tournamentsArrays.length, 1);
  selectedRange.insertCheckboxes();
  activeRange.insertCheckboxes();
}

function readTournamentDataFromSheet(sheet, currentOrder){
  let firstCol = TOURNAMENTS_SECTIONS_FIRST_COLUMN + TOURNAMENTS_SECTIONS_STEP*currentOrder;
  let firstRow = TOURNAMENTS_SECTIONS_FIRST_ROW;
  //let gamesCount = Number(sheet.getRange(firstRow + TOURNAMENT_SECTION_HEADERS.indexOf('Games Number'), firstCol + 1, 1, 1).getValue());
  // headers + padding + title + games number
  return sheet.getSheetValues(firstRow, firstCol, -1 /*TOURNAMENT_SECTION_HEADERS.length + GAMES_SECTIONS_PADDING + 1 + gamesCount*/, GAME_HEADERS.length);
}

function processGame(game){
  let first_team = game['CompetitorA']
  let second_team = game['CompetitorB']
  return {
    teamACode: first_team['TeamCode'], 
    teamBCode: second_team['TeamCode'], 
    gameDate: new Date(game['StartTimeUtc']).toISOString(), 
    teamAScore: first_team['Score'], 
    teamBScore: second_team['Score']
  };
}

function readTournamentDataFromFIBA(tournament){
  var url =`https://livecache.sportresult.com/node/db/FIBASTATS_PROD/${tournament.tournamentId}_SCHEDULELS_JSON.json`;
  // console.log('fetching ' + url);
  var result = UrlFetchApp.fetch(url);
  var response = result.getContentText();
  response = JSON.parse(response)['content'];
  let games = Object.keys(response['full']['Games']).map(k => processGame(response['full']['Games'][k]));
  let headersValues = TOURNAMENT_SECTION_HEADERS.map(h => [h.displayName, h.key ? tournament[h.key] : '']);
  let gamesValues = games.map(g => GAME_HEADERS.map(h => h.key ? g[h.key] : ''));
  let sectionValues = [...headersValues, ..._.range(GAMES_SECTIONS_PADDING).map(r => []), [GAMES_SECTIONS_TITLE], GAME_HEADERS.map(h => h.displayName), ...gamesValues];
  return sectionValues;
}

function fetchTournamentDataToWrite(sheet, tournament, currentTournamentsOrder){
  if(tournament.active || !currentTournamentsOrder[tournament.tournamentId]){
    let padded = readTournamentDataFromFIBA(tournament).map(r => _.assign(_.fill(new Array(GAME_HEADERS.length), ''), r));
    return padded;
  }
  return readTournamentDataFromSheet(sheet, currentTournamentsOrder[tournament.tournamentId].order);
}

function writeTournament(sheet, tournamentValues, order){
  var rules = sheet.getConditionalFormatRules();
  let firstSectionRow = TOURNAMENTS_SECTIONS_FIRST_ROW;
  let firstSectionColumn = TOURNAMENTS_SECTIONS_FIRST_COLUMN + order * TOURNAMENTS_SECTIONS_STEP;
  let valuesRange = sheet.getRange(firstSectionRow, firstSectionColumn, tournamentValues.length, tournamentValues[0].length);
  valuesRange.setValues(tournamentValues);

  let sectionHeadersRange = sheet.getRange(firstSectionRow, firstSectionColumn, GAMES_SECTIONS_OFFSET_FROM_SECTION, 1);
  let gamesHeadersRange = sheet.getRange(firstSectionRow + GAMES_SECTIONS_OFFSET_FROM_SECTION, firstSectionColumn, 1, GAME_HEADERS.length);
  sectionHeadersRange.setFontWeight("bold");
  gamesHeadersRange.setFontWeight("bold");

  let firstGameRow = firstSectionRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
  let firstTeamACodeColumn = firstSectionColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM A');
  let firstTeamBCodeColumn = firstSectionColumn + GAME_HEADERS.findIndex(h => h.displayName == 'TEAM B');
  let firstWinnerColumn = firstSectionColumn + GAME_HEADERS.findIndex(h => h.displayName == 'Winner');

  let firstTeamACodeCell = sheet.getRange(firstGameRow, firstTeamACodeColumn, 1, 1).getA1Notation();
  let firstTeamBCodeCell = sheet.getRange(firstGameRow, firstTeamBCodeColumn, 1, 1).getA1Notation();
  let firstWinnerCell = sheet.getRange(firstGameRow, firstWinnerColumn, 1, 1).getA1Notation();

  let teamACodeRange = sheet.getRange(firstGameRow, firstTeamACodeColumn, tournamentValues.length - GAMES_SECTIONS_OFFSET_FROM_SECTION - 1, 1);
  let teamBCodeRange = sheet.getRange(firstGameRow, firstTeamBCodeColumn, tournamentValues.length - GAMES_SECTIONS_OFFSET_FROM_SECTION - 1, 1);
  let teamARule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=and(not(isblank(${firstTeamACodeCell})),EQ(${firstTeamACodeCell}, ${firstWinnerCell}))`)
    .setBackground('green')
    .setRanges([teamACodeRange])
    .build();
  let teamBRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=and(not(isblank(${firstTeamBCodeCell})),EQ(${firstTeamBCodeCell}, ${firstWinnerCell}))`)
    .setBackground('green')
    .setRanges([teamBCodeRange])
    .build();
  rules.push(teamARule);
  rules.push(teamBRule);
  // apply tournaments formulas
  for(let [i, h] of TOURNAMENT_SECTION_HEADERS.entries()){
    if(h.formula){
      let val = h.formula(sheet, firstSectionRow, firstSectionColumn);
      let hRow = firstSectionRow + i;
      let hColumn = firstSectionColumn + 1;
      let hCell = sheet.getRange(hRow, hColumn, 1, 1);
      hCell.setFormula(val);
    }
  }

  // apply games formulas
  for(let [i, h] of GAME_HEADERS.entries()){
    if(h.formula){
      let val = h.formula(sheet, firstSectionRow, firstSectionColumn, 0);
      let gameRow = firstSectionRow + GAMES_SECTIONS_OFFSET_FROM_SECTION + 1;
      let gamesCount = tournamentValues.length - gameRow + 1;
      let hColumn = firstSectionColumn + i;
      let hCell = sheet.getRange(gameRow, hColumn, 1, 1);
      hCell.setFormula(val);
      let gamesCells = sheet.getRange(gameRow, hColumn, gamesCount, 1);
      hCell.autoFill(gamesCells, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
  }
  sheet.setConditionalFormatRules(rules);
}

function updateCompetition(spreadsheet, competitionName, competitionTournaments){
  let competitionSheet = ensureSheetExist(spreadsheet, competitionName);
  // competitionTournaments = _.sortBy(competitionTournaments, ['order']);
  let idsRange = getRowToEnd(competitionSheet.getRange(TOURNAMENTS_SECTIONS_FIRST_ROW + TOURNAMENT_SECTION_HEADERS.findIndex(h => h.displayName == 'ID'), TOURNAMENTS_SECTIONS_FIRST_COLUMN + 1, 1, 2).getA1Notation());
  let currentIds = competitionSheet.getRange(idsRange).getValues()[0]
    .map((tid, index) => ({id: tid, order: index % 9 == 0 && Number(tid) > 0 ? index / 9 : -1})).filter(o => o.order >= 0);
  let currentTournamentsIndexes = _.keyBy(currentIds, 'id');
  let tournamentsValuesToWrite = competitionTournaments.map(t => 
    ({
      id: t.tournamentId,
      values: fetchTournamentDataToWrite(competitionSheet, t, currentTournamentsIndexes),
      order: t.order}));
  competitionSheet.getRange(1, 1, competitionSheet.getMaxRows(), competitionSheet.getMaxColumns()).clearDataValidations();
  competitionSheet.clear();
  for(let [order, tournamentToWrite] of tournamentsValuesToWrite.entries()){
    writeTournament(competitionSheet, tournamentToWrite.values, order);
  }
}

function updateCompetitions(spreadsheet, competitions){
  for(let competition of Object.keys(competitions)){
    updateCompetition(spreadsheet, competition, competitions[competition]);
  }
}

function updateMainSheet(spreadsheet, competitions){
  let mainSheet = ensureSheetExist(spreadsheet, 'Main');
  mainSheet.getRange(1, 1, mainSheet.getMaxRows(), mainSheet.getMaxColumns()).clearDataValidations();
  mainSheet.clear();
  var currentCol = MAIN_SECTIONS_FIRST_COLUMN;
  for(let competitionName of Object.keys(competitions)){
    let competitionSheet = spreadsheet.getSheetByName(competitionName);
    let competitionCell = mainSheet.getRange(MAIN_SECTIONS_FIRST_ROW, currentCol, 1, 1);
    competitionCell.setValue(competitionName);
    competitionCell.setFontWeight('bold');
    let competitionHeadersRange = mainSheet.getRange(MAIN_SECTIONS_FIRST_ROW + MAIN_TOURNAMENTS_SECTIONS_PADDING + 1, currentCol, 1, MAIN_TOURNAMENTS_SECTION_HEADERS.length);
    competitionHeadersRange.setValues([MAIN_TOURNAMENTS_SECTION_HEADERS.map(h => h.displayName)]);
    competitionHeadersRange.setFontWeight('bold');
    for(let [i, tourObj] of competitions[competitionName].entries()){
      for(let [headerIndex, h] of MAIN_TOURNAMENTS_SECTION_HEADERS.entries()){
        // let tournamentTitleMainRange = mainSheet.getRange(MAIN_SECTIONS_FIRST_ROW + MAIN_TOURNAMENTS_SECTION_HEADERS.length + MAIN_TOURNAMENTS_SECTIONS_PADDING + 1 + i, currentCol, 1, 1);
        // tournamentTitleMainRange.setValue(tourObj.desc);
        let headerCell = mainSheet.getRange(MAIN_SECTIONS_FIRST_ROW + MAIN_TOURNAMENTS_SECTIONS_PADDING + 2 + i, currentCol + headerIndex, 1, 1);
        let headerCompCellColumn = TOURNAMENTS_SECTIONS_FIRST_COLUMN + i * TOURNAMENTS_SECTIONS_STEP + 1;
        let headerCompCellRow = TOURNAMENTS_SECTIONS_FIRST_ROW + TOURNAMENT_SECTION_HEADERS.findIndex(compH => compH.displayName == h.tournamentHeaderDisplayName);
        let headerCompCell = competitionSheet.getRange(headerCompCellRow, headerCompCellColumn, 1, 1).getA1Notation();
        headerCell.setFormula(`='${competitionName}'!${headerCompCell}`);
      }
    }
    currentCol += MAIN_SECTIONS_STEP;
  }
}

function resetData(){
  refreshData(true);
}

function refreshData(reset=false){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let tournamentsConfig = getCurrentTournamentsConfig(reset);
  let selectedTournaments = _.sortBy(_.values(tournamentsConfig).filter(t => t.selected), ['order']);
  let competitions = _.groupBy(selectedTournaments, (t) => t.competition != '' ? t.competition : 'Uncategorized');
  updateCompetitions(spreadsheet, competitions);
  updateMainSheet(spreadsheet, competitions);
  let controllerSheet = spreadsheet.getSheetByName('Controller');
  let cellToUpdate = controllerSheet.getRange(reset ? CONTROLLER_LAST_RESET_CELL : CONTROLLER_LAST_REFRESH_CELL);
  cellToUpdate.setValue((new Date()).toISOString());
}
