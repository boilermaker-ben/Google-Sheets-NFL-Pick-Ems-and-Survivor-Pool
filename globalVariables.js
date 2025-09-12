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

const LEAGUE = 'NFL';
const TEAMS = 32;
const REGULAR_SEASON = 18; // Regular season matchups
const WEEKS = 23; // Total season weeks (including playoffs)
const WEEKS_TO_EXCLUDE = [22]; // Break before Superbowl
const MAXGAMES = TEAMS/2;
const SCOREBOARD = "https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard";

const DAY = {  0: {"name":"Sunday","index":0},  1: {"name":"Monday","index":1},  2: {"name":"Tuesday","index":2},  3: {"name":"Wednesday","index":-4},  4: {"name":"Thursday","index":-3},  5: {"name":"Friday","index":-2},  6: {"name":"Saturday","index":-1} };

const WEEKNAME = { 19: {"name":"WildCard","teams":12,"matchups":6}, 20: {"name":"Divisional","teams":8,"matchups":4}, 21: {"name":"Conference","teams":4,"matchups":2}, 23: {"name":"SuperBowl","teams":2,"matchups":1} };

const LEAGUE_DATA = {
  "ARI": {
    "division": "NFC West",
    "division_opponents": ["LAR", "SEA", "SF"],
    "colors": [
      "#97233F",
      "#000000",
      "#FFB612",
      "#FFFFFF"
    ],
    "mascot": "🐦",
    "colors_emoji": "🔴⚫"
  },
  "ATL": {
    "division": "NFC South",
    "division_opponents": ["CAR", "NO", "TB"],
    "colors": [
      "#010101",
      "#A6192E",
      "#FFFFFF",
      "#B2B4B2"
    ],
    "mascot": "🦜",
    "colors_emoji": "⚫🔴"
  },
  "BAL": {
    "division": "AFC North",
    "division_opponents": ["CIN", "CLE", "PIT"],
    "colors": [
      "#24125F",
      "#FFFFFF",
      "#9A7611",
      "#010101"
    ],
    "mascot": "🐦‍⬛",
    "colors_emoji": "🟣🟡"
  },
  "BUF": {
    "division": "AFC East",
    "division_opponents": ["MIA", "NE", "NYJ"],
    "colors": [
      "#003087",
      "#C8102E",
      "#FFFFFF",
      "#091F2C"
    ],
    "mascot": "🐃",
    "colors_emoji": "🔵🔴"
  },
  "CAR": {
    "division": "NFC South",
    "division_opponents": ["ATL", "NO", "TB"],
    "colors": [
      "#101820",
      "#0085CA",
      "#B2B4B2",
      "#FFFFFF"
    ],
    "mascot": "🐈‍⬛",
    "colors_emoji": "⚫🔵"
  },
  "CHI": {
    "division": "NFC North",
    "division_opponents": ["DET", "GB", "MIN"],
    "colors": [
      "#091F2C",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐻",
    "colors_emoji": "🔵🟠"
  },
  "CIN": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CLE", "PIT"],
    "colors": [
      "#010101",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐅",
    "colors_emoji": "⚫🟠"
  },
  "CLE": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "PIT"],
    "colors": [
      "#311D00",
      "#EB3300",
      "#FFFFFF",
      "#EDC8A3"
    ],
    "mascot": "🟠",
    "colors_emoji": "🟤🟠"
  },
  "DAL": {
    "division": "NFC East",
    "division_opponents": ["NYG", "PHI", "WSH"],
    "colors": [
      "#0C2340",
      "#FFFFFF",
      "#87909A",
      "#7F9695"
    ],
    "mascot": "🤠",
    "colors_emoji": "🔵⚪"
  },
  "DEN": {
    "division": "AFC West",
    "division_opponents": ["KC", "LAC", "LV"],
    "colors": [
      "#0C2340",
      "#FC4C02",
      "#FFFFFF"
    ],
    "mascot": "🐴",
    "colors_emoji": "🔵🟠"
  },
  "DET": {
    "division": "NFC North",
    "division_opponents": ["CHI", "GB", "MIN"],
    "colors": [
      "#0069B1",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "🦁",
    "colors_emoji": "⚪🔵"
  },
  "GB": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "MIN"],
    "colors": [
      "#183029",
      "#FFB81C",
      "#FFFFFF"
    ],
    "mascot": "🧀",
    "colors_emoji": "🟢🟡"
  },
  "HOU": {
    "division": "AFC South",
    "division_opponents": ["IND", "JAX", "TEN"],
    "colors": [
      "#1D1F2A",
      "#E4002B",
      "#FFFFFF",
      "#0072CE"
    ],
    "mascot": "🐂",
    "colors_emoji": "🔴🔵"
  },
  "IND": {
    "division": "AFC South",
    "division_opponents": ["HOU", "JAX", "TEN"],
    "colors": [
      "#003A70",
      "#FFFFFF",
      "#A2AAAD",
      "#1D252D"
    ],
    "mascot": "🐎",
    "colors_emoji": "🔵⚪"
  },
  "JAX": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "TEN"],
    "colors": [
      "#006271",
      "#D29F13",
      "#010101",
      "#9A7611"
    ],
    "mascot": "🌴",
    "colors_emoji": "🟡🔵"
  },
  "KC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "LAC", "LV"],
    "colors": [
      "#C8102E",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🏹",
    "colors_emoji": "🔴🟡"
  },
  "LAC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LV"],
    "colors": [
      "#0072CE",
      "#FFB81C",
      "#FFFFFF",
      "#0C2340"
    ],
    "mascot": "⚡",
    "colors_emoji": "🔵🟡"
  },
  "LAR": {
    "division": "NFC West",
    "division_opponents": ["ARI", "SEA", "SF"],
    "colors": [
      "#1E22AA",
      "#FFD100",
      "#D7D2CB",
      "#FFFFFF"
    ],
    "mascot": "🐏",
    "colors_emoji": "🔵🟡"
  },
  "LV": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LAC"],
    "colors": [
      "#010101",
      "#A2AAAD",
      "#FFFFFF",
      "#87909A"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "⚫⚪"
  },
  "MIA": {
    "division": "AFC East",
    "division_opponents": ["BUF", "NE", "NYJ"],
    "colors": [
      "#008C95",
      "#FC4C02",
      "#FFFFFF",
      "#005776"
    ],
    "mascot": "🐬",
    "colors_emoji": "🟡🔵"
  },
  "MIN": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "GB"],
    "colors": [
      "#582C83",
      "#FFC72C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⚔️",
    "colors_emoji": "🟣🟡"
  },
  "NE": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NYJ"],
    "colors": [
      "#0C2340",
      "#C8102E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🥁", //🧦
    "colors_emoji": "🔵🔴"
  },
  "NO": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "TB"],
    "colors": [
      "#010101",
      "#D3BC8D",
      "#FFFFFF",
      "#A28D5B"
    ],
    "mascot": "⚜️",
    "colors_emoji": "⚫🟡"
  },
  "NYG": {
    "division": "NFC East",
    "division_opponents": ["DAL", "PHI", "WSH"],
    "colors": [
      "#001E62",
      "#A6192E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🏗️",
    "colors_emoji": "🔵🔴"
  },
  "NYJ": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NE"],
    "colors": [
      "#115740",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "✈️",
    "colors_emoji": "🟢⚪"
  },
  "PHI": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "WSH"],
    "colors": [
      "#004851",
      "#E3E5E6",
      "#545859",
      "#010101"
    ],
    "mascot": "🦅",
    "colors_emoji": "🟢⚫"
  },
  "PIT": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "CLE"],
    "colors": [
      "#010101",
      "#FFB81C",
      "#FFFFFF",
      "#C8102E"
    ],
    "mascot": "🏭",
    "colors_emoji": "⚫🟡"
  },
  "SEA": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SF"],
    "colors": [
      "#0C2340",
      "#78BE21",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🌊",
    "colors_emoji": "🔵🟢"
  },
  "SF": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SEA"],
    "colors": [
      "#A6192E",
      "#B9975B",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⛏️",
    "colors_emoji": "🔴🟡"
  },
  "TB": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "NO"],
    "colors": [
      "#010101",
      "#A6192E",
      "#3D3935",
      "#DC4405"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "🔴⚫"
  },
  "TEN": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "JAX"],
    "colors": [
      "#0C2340",
      "#418FDE",
      "#B2B4B2",
      "#C8102E"
    ],
    "mascot": "🛡️",
    "colors_emoji": "🔵🔴"
  },
  "WSH": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "PHI"],
    "colors": [
      "#651C32",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🎖️",
    "colors_emoji": "🟤🟡"
  }
};
