var ss = SpreadsheetApp.getActiveSpreadsheet();
var theme = ss.getSpreadsheetTheme();
var ui = SpreadsheetApp.getUi();

function doShit(uploading) {
  var sheet = ss.getSheetByName("Editor")
  
  let objRegex = /obj\/(\d+)(-(\d+))?\.png/
  let range = sheet.getRange(1,sheet.getFrozenColumns() + 1, sheet.getLastRow(), sheet.getLastColumn())
  let cell = range.getFormulas()
  let text = range.getValues()
  let bgs = range.getBackgroundObjects()
  let sizes = range.getFontSizes()
  
  let levelData = ""

  for (var r in cell) {
    for (var c in cell[r]) {
      
      let xPos = 15 + (c * 30)
      let yPos = 15 + ((49-r) * 30)
      let val = cell[r][c]
      let txt = text[r][c]
      
      if (val && val.match(objRegex)[1]) {
        
        let data = val.match(objRegex)
        let obj = data[1]
        let rotation = data[3]
        let extraData = ""
        
        let offset = offsets[obj]
        if (offset) {
          
          let off = !Array.isArray(offset) ? [0, offset] : offset
          off = [off[0], off[1], off[2] || 0, off[3] || 0]
          
          switch(+rotation || 0) {
            case 0: xPos += off[0]; yPos += off[1]; break;
            case 90: xPos += off[1]; yPos -= off[0]; xPos += off[2] || 0; break;
            case 180: xPos -= off[0]; yPos -= off[1]; xPos += off[3] || 0; yPos -= off[2] || 0; break;
            case 270: xPos -= off[1]; yPos += off[0]; yPos -= off[3]; break;
          }
          
        }
        
        if (+rotation == 315) {
          rotation = "-45"
        }
        
        if (speedPortals.includes(+obj)) {  // speed portal checkbox
          extraData = ',13,1'
        }
        
        else if (Object.keys(triggers).includes(obj)) {  // triggers
          let bgColor = bgCol(bgs[r][c])
          let fontSize = sizes[r][c]
          let fadeTime = fontSize / 10
          if (fadeTime == 0.1) fadeTime = 0
          extraData = `,7,${bgColor.r},8,${bgColor.g},9,${bgColor.b},10,${fadeTime},35,1,23,${triggers[obj]}`
        }
        
        levelData += `1,${obj},2,${xPos},3,${yPos}${rotation ? `,6,${rotation}` : ""}${extraData};`
      }
      
      else if (txt) { // text object
        levelData += `1,914,2,${xPos},3,${yPos},31,${Utilities.base64EncodeWebSafe(txt)};` 
      }
      
    }
  }
  
  if (!uploading) {
      let settings = levelInfo()
      saveLevelData(settings.str + levelData)
  }
  
  return levelData
  
}

function saveLevelData(levelCode, prompt=true) {
  if (levelCode.length >= 49900) {
    levelCode = gzip(levelCode)
    if (prompt && levelCode.length >= 49900) return ui.alert('Save error!', 'The level data is too large to be saved in a single spreadsheet sell, even after compressing!', ui.ButtonSet.OK) 
    else if (prompt) ui.alert('Level data saved!', 'The level data has been saved in the "Data" tab, but has been compressed with GZIP to stay below the character limit.', ui.ButtonSet.OK) 
  }
  else if (prompt) ui.alert('Level data saved!', 'The level data has been saved in the "Data" tab.', ui.ButtonSet.OK) 
  return ss.getSheetByName("Data").getRange('A1').setValue(levelCode);
}

function levelInfo() {
  
  var sheet = ss.getSheetByName("Settings")
  let info = {}
  let kS38 = "kS38,"
  let range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
  let settings = range.getValues() 
  let bgs = range.getBackgroundObjects()
  settings.forEach((x, y) => {
    if (x[2].startsWith("kS38")) {
      let c = bgCol(bgs[y][1])
      let channel = x[2].split("/")[1]
      kS38 += `1_${c.r}_2_${c.g}_3_${c.b}_4_-1_6_${channel}_7_1|`
    }
    else if (x[0].length) info[x[2]] = x[1]                  
  })
  info.str = kS38 + `,kA13,${info.kA13},kA15,${getVal(info.kA15)},kA16,${getVal(info.kA16)},kA14,,kA6,${getVal(info.kA6)},kA7,${getVal(info.kA7)},kA17,${getVal(info.kA17)},kA18,${getVal(info.kA18, true)},kS39,0,kA2,${getVal(info.kA2, true)},kA3,${getVal(info.kA3)},kA8,${getVal(info.kA8)},kA4,${getVal(info.kA4, true)},kA9,0,kA10,${getVal(info.kA10)},kA11,0;`
  return info
}

function uploadLevel() {
  
  let settings = levelInfo()
    
  let lvlString = doShit(true)
  let objCount = (lvlString.split(";").length) - 1
  let coinCount = Math.min(lvlString.split(";1,1329").length -1, 3)
  saveLevelData(settings.str + lvlString, false)
  let levelData = gzip(settings.str + lvlString)
  
  let prompt1 = ui.prompt('Upload level (1/2)', 'Please enter your GD username:', ui.ButtonSet.OK);  
  let username = prompt1.getResponseText()
  if (prompt1.getSelectedButton() == ui.Button.CLOSE) return
  
  let prompt2 = ui.prompt('Upload level (2/2)', 'Please enter your GD password:', ui.ButtonSet.OK); 
  if (prompt2.getSelectedButton() == ui.Button.CLOSE) return
  let password = prompt2.getResponseText()
  
  let loginParams = `userName=${username}&password=${password}&udid=69&secret=Wmfv3899gc9`
  let login = UrlFetchApp.fetch('http://boomlings.com/database/accounts/loginGJAccount.php', { method : 'post', contentType: 'application/x-www-form-urlencoded', payload: loginParams})
  let loginRes = login.getContentText()
  if (login.getResponseCode() != 200 || !loginRes || loginRes == "-1") return ui.alert('Invalid login!', 'Double check your username/password, or try again later.', ui.ButtonSet.OK) 
  let accID = loginRes.split(",")[0]
  
  let GJP = XOR(password, '37526')
  let seed2 = ""
  
  var params = {
    levelString: levelData,
    levelName: settings.levelName,
    levelDesc: Utilities.base64EncodeWebSafe(settings.levelDesc),
    userName: username,
    accountID: accID,
    gjp: GJP,
    levelVersion: +settings.levelVersion,
    levelID: 0,
    levelLength: getVal(settings.levelLength, true),
    audioTrack: getVal(settings.audioTrack, true),
    songID: +settings.songID,
    auto: 0,
    password: +settings.password,
    original: 0,
    twoPlayer: 0,
    unlisted: getVal(settings.unlisted),
    ldm: 0,
    objects: objCount,
    coins: coinCount,
    requestedStars: +settings.requestedStars,
    gameVersion: 21,
    binaryVersion: 35,
    extraString: "1_7_7_0_1_3",
    secret: "Wmfd2893gb7"
  };
  
  let spacing = Math.floor(levelData.length / 50)
  for (i=0; i < 50; i++) {
    seed2 += levelData[spacing*i]
  }
  
  seed2 = sha1(seed2 + "xI25fpAapCQg")
  seed2 = XOR(seed2, '41274')
  params.seed2 = seed2
 
  let options = {
    method : 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload : Object.keys(params).map(k => k + '=' + params[k]).join('&')
  };
  
  let upload = UrlFetchApp.fetch('http://boomlings.com/database/uploadGJLevel21.php', options);
  let uploadRes = upload.getContentText()
  if (upload.getResponseCode() != 200 || !uploadRes || uploadRes == "-1") return ui.alert('Upload error!', 'Your level could not be uploaded to the GD servers.\nRemove any illegal values, or try again later', ui.ButtonSet.OK) 
  else return ui.alert('Upload complete!', `Your level was successfully uploaded!${getVal(settings.unlisted) ? " (unlisted)" : ""}\nID: ${uploadRes}`, ui.ButtonSet.OK) 
}

function getVal(str, bump) {
  if (str == "Yes" || str == "No") return str == "Yes" ? 1 : 0
  let num = Number(str.split(" ")[0])
  if (bump) num -= 1
  return num
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('GD')
  .addItem('Generate level data', 'doShit')
  .addItem('Upload level', 'uploadLevel')
  .addToUi();
};

function bgCol(col) {
  let type = col.getColorType()
  if (type == "THEME") {  // man
    let themeType = col.asThemeColor().getThemeColorType()
    col = theme.getConcreteColor(themeType)
  }
  let rgb = col.asRgbColor()
  return {r: rgb.getRed(), g: rgb.getGreen(), b: rgb.getBlue()}
}


/*
=================
CRYPTO STUFF (ew)
=================
*/

function sha1(text) {
  let res = ""
  let sha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, text)
  sha.forEach(x => {
	 if (x < 0) x += 256
	 if (x.toString(16).length == 1) res += '0'
	 res += x.toString(16)
  })
  return res
}

function XOR(str, key) {
  
  // .fromCodePoint() acts REALLY weird on Google scripts, but un-nesting the code seems to fix it 
  let step1 = str.split('').map((char, i) => char.charCodeAt(0) ^ key.toString().charCodeAt(i % key.toString().length))
  let step2 = String.fromCodePoint(...step1)
  return Utilities.base64EncodeWebSafe(step1)
  
}

function gzip(str) {
  let zip = Utilities.gzip(Utilities.newBlob(str))
  return Utilities.base64EncodeWebSafe(zip.getBytes());
}


// JSON stuff, I guess

let speedPortals = [200, 201, 202, 203, 1334]
let triggers = {29: 1000, 30: 1001, 105: 1004, 900: 1009, 915: 1002}

const offsets = {
    "40": 8,    // half block
    "62": 7,    // wavy block
    "64": [-8, 8], // wavy corner
    "468": 14.25,  // outline
  
    "39": -8.5, // half spike
    "178": -8.5, // half electro spike
  
    "9": -12, // spike floor
    "61": -11, // wavy spike floor
    "243": -11, // wavy spikes, right end
    "244": -11, // wavy spikes, left end
  
    "157": -2, // wavy deco, center
    "158": -2, // wavy deco, left
    "159": -2, // wavy deco, left
  
    "130": [45, 0], // 4 wide cloud
    "129": [30, 0], // 3 wide cloud
  
    "35": -13,   // yellow pad
    "67": -13,   // blue pad
    "140": -13,  // pink pad
    "1332": -13, // red pad
  
    "10": [0, -30, 60],  // blue gravity portal
    "11": [0, -30, 60],  // yellow gravity portal
    "99": [0, -30, 60],  // large portal
    "101": [0, -30, 60], // mini portal
    "201": [0, -30, 60],  // 1x portal
    "286": [0, -30, 60], // single portal
    "287": [0, -30, 60], // dual portal
  
    "45": [15, -30, 60, 30],  // mirror portal
    "46": [15, -30, 60, 30],  // un-mirror portal
    "200": [15, -30, 60, 30],  // -1x portal
    "202": [15, -30, 60, 30],  // 2x portal
  
    "12": [0, -30, 60],   // cube portal
    "13": [0, -30, 60],   // ship portal
    "47": [0, -30, 60],   // ball portal
    "111": [0, -30, 60],  // ufo portal
    "660": [0, -30, 60],  // wave portal
    "745": [0, -30, 60],  // robot portal
    "1331": [0, -30, 60], // spider portal
  
    "203": [30, -15, 30, 60],  // 3x portal
    "1334": [30, -15, 30, 60], // 4x portal
  
    "16": -1,  // medium antenna 
  
    "85": [30, -30], // large sawblade deco
    "88": [30, -30], // large sawblade
  
    "41": [0, -10, 30],  // chain
    "106": [0, -12, 30], // fancy chain
  
    "18": [44, 4, 0, 90],  // 4 wide spike deco
    "19": [30, 4, 0, 60],  // 3 wide spike deco
    "20": [16, -2, 0, 30], // 2 wide spike deco
    "21": -8,   // 1 wide spike deco
}
