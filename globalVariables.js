//------------------------------------------------------------------------
// GLOBAL VARIABLES - for easy modification in the future
const league = 'NFL'; // Hopefully I'll be able to support NCAAF at some point
const fbTeams = 32;
const maxGames = fbTeams/2;
const weeklySheetPrefix = 'WK';
const scoreboard = 
    league == 'NFL' ? 'https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard' :
    (league == 'NCAAF' ? 'https://site.api.espn.com/apis/site/v2/sports/football/college-football/scoreboard' : null);
const schedulePrefix = 'https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/';
const scheduleSuffix = '?view=proTeamSchedules';
const fallbackYear = 2024;
const dayColors = ['#fffdcc','#e7fed1','#cffdda','#bbfbe7','#adf7f5'];
const dayColorsFilled = ['#fffb95','#d4ffa6','#abffbf','#89fddb','#74f7f3'];
const configTabColor = '#ff9561';
const generalTabColor = '#aaaaaa';
const winnersTabColor = '#ffee00';
const scheduleTabColor = '#472a24';
