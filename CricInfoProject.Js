//The purpose of this project is to extract information of worldcup 2019 from cricinfo
//and present that in the form of excel & pdf scoreCards
// the real purpose of this project is to learn how to extract information using JS
//FUN PROJECT
//A very good reason to ever make a project is to have good fun


//npm init
//npm install minimist
//npm install axios
//npm install jsdom
//npm install excel4node
//npm pdf-lib


// node CricInfoProject.Js  --excel WorldCupData.csv --dataFolder data --source https://www.espncricinfo.com/series/icc-cricket-world-cup-2019-1144415/match-results

let minimist = require("minimist");
let axios = require("axios");
let jsdom = require("jsdom");
let excel = require("excel4node");
let pdf = require("pdf-lib");
let fs = require("fs");
let args = minimist(process.argv);

//Download using axios
//Read using JSDOM
//Make Excel using excel4node
//make pdf using pdf-lib

let responseKaPromise = axios.get(args.source);
responseKaPromise.then(function(response){
    let html = response.data;
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;
    let matches = [];
    let matchScoreDivs = document.querySelectorAll("div.match-score-block");
    for(let i = 0 ; i<matchScoreDivs.length; i++){
        let match = {

        };
        //Team Names
        let namesPara = matchScoreDivs[i].querySelectorAll("p.name");
        match.t1 = namesPara[0].textContent;
        match.t2 = namesPara[1].textContent;

        //Team Scores
        let spanScore = matchScoreDivs[i].querySelectorAll("div.score-detail > span.score");
        match.t1s = "";
        match.t2s = "";
        if(spanScore.length==2){
            match.t1s = spanScore[0].textContent;
            match.t2s = spanScore[1].textContent;
        }else if(spanScore.length == 1){
            match.t1s = spanScore[0].textContent;
            match.t2s = "";
        }else{
            match.t1s = "";
            match.t2s = "";
        }

        //Result
        let resultDiv = matchScoreDivs[i].querySelector("div.status-text > span");
        match.result = resultDiv.textContent;

        //FLAGS
        let flags = matchScoreDivs[i].querySelectorAll(".team .name-detail .img.flag");
        let tumpu = JSON.stringify(flags);
        console.log(flags.textContent);
        //Pushing all team matches blocks in our array
        matches.push(match);
    }
    let matchesKaJason = JSON.stringify(matches);
    fs.writeFileSync("matches.json", matchesKaJason, "utf-8");

    let teams = [];
    for(let i = 0 ; i<matches.length; i++){
        addTeamNamesOnly(teams,matches[i].t1);
        addTeamNamesOnly(teams,matches[i].t2);
    }

    for(let i = 0 ; i<matches.length ; i++){
        addTeamDetails(teams, matches[i].t1, matches[i].t2,matches[i].t1s,matches[i].t2s,matches[i].result);
        addTeamDetails(teams, matches[i].t2, matches[i].t1,matches[i].t2s,matches[i].t1s,matches[i].result);
    }

    let teamKaJason = JSON.stringify(teams);
    fs.writeFileSync("teams.json", teamKaJason, "utf-8");   
    
    //Create Excel File
    let wb = new excel.Workbook();
    for(let i = 0 ; i<teams.length ; i++){
        let sheet = wb.addWorksheet(teams[i].name);
        sheet.cell(1,1).string("Opponent");
        sheet.cell(1,2).string("TeamScore");
        sheet.cell(1,3).string("Opp Score");
        sheet.cell(1,4).string("Result");

        for(let j = 0 ; j<teams[i].matches.length ; j++){
            let vs = teams[i].matches[j].vs;
            sheet.cell(j+2, 1).string(vs);
            let teamScore = teams[i].matches[j].ourScore;
            sheet.cell(j+2, 2).string(teamScore);
            let OppScore = teams[i].matches[j].enemyScore;
            sheet.cell(j+2,3).string(OppScore);
            let result = teams[i].matches[j].matchResult;
            sheet.cell(j+2, 4).string(result);
        }
    }
    wb.write(args.excel);
    let path = require("path");
    //Creating Folders
    for(let i = 0 ; i<teams.length ; i++){
        let folderName = path.join("WorldCup", teams[i].name);
        fs.mkdirSync(folderName);
        
        // Creating PDF's
        for(let j = 0 ; j<teams[i].matches.length; j++){
            let matchFileName = path.join(folderName,teams[i].matches[j].vs+".pdf");
            createScoreCard(teams[i].name, teams[i].matches[j],matchFileName);
        }
    }
}).catch(function(err){
    console.log(err);
})
function addTeamNamesOnly(teams, teamName){
    let idx = -1;
    for(let i  = 0 ; i<teams.length ; i++){
        if(teams[i].name == teamName){
            idx = i;
            break;
        }
    }
    if(idx == -1){
        let team = {
            name : teamName,
            matches:[]
        }
        teams.push(team);
    }
}
function addTeamDetails(teams,teamName, oppName, teamScore, oppScore, result){
    let idx = -1;
    for(let i  = 0 ; i<teams.length ; i++){
        if(teams[i].name == teamName){
            idx = i;
            break;
        }
    }

    teams[idx].matches.push({
        vs : oppName,
        ourScore : teamScore,
        enemyScore : oppScore,
        matchResult : result
    })
}
function createScoreCard(teamName, matches, matchFileName){
    let t1 = teamName;
    let t2 = matches.vs;
    let t1s = matches.ourScore;
    let t2s = matches.enemyScore;
    let result = matches.matchResult;

    let originalBytes = fs.readFileSync("temp.pdf");
    let pdfKaPromise = pdf.PDFDocument.load(originalBytes);
    pdfKaPromise.then(function(pdfDoc){
        let page = pdfDoc.getPage(0);
        page.drawText(t1,{
            x: 45,
            y: 250,
            size: 25
        });

        page.drawText(t2,{
            x: 45 ,
            y: 190,
            size: 25
        });

        page.drawText(t1s,{
            x: 290 ,
            y: 250,
            size: 20
        });

        page.drawText(t2s,{
            x: 290 ,
            y: 190,
            size: 20
        });

        page.drawText(result,{
            x: 40,
            y: 85,
            size: 20
        });

        let promiseToSavePdf = pdfDoc.save();
        promiseToSavePdf.then(function(changedBytes){
            fs.writeFileSync(matchFileName,changedBytes);
        }).catch(function(err){
            console.log(err);
        })
    }).catch(function(err){
        console.log(err);
    })
}
