const xlsxFile = require("read-excel-file/node");
const Excel = require("exceljs");
const Teams = require("./teams");
const teams = Teams;
let scores = [];
const fs = require("firebase-admin");
async function getTotal() {
  await xlsxFile("./data.xlsx").then((rows) => {
    scores = rows;
    rows.shift();
    let result = new Map(
      rows.map((row) => {
        let email = row[0].toLowerCase();
        row.shift();
        return [email, row];
      })
    );
    teams.map((team) => {
      let arr = [];
      team.members.map((member) => {
        if (result.get(member.toLowerCase()) != undefined) {
          arr.push(result.get(member.toLowerCase()));
        }
      });
      let maxScores = [];

      if (arr[0] != undefined) {
        for (let i = 0; i < arr[0].length; i++) {
          let max = 0;
          for (let j = 0; j < arr.length; j++) {
            if (arr[j][i] >= max) {
              max = arr[j][i];
            }
          }
          maxScores.push(max);
        }
        const totScore = maxScores.reduce((a, b) => a + b, 0);
        team.score = totScore;
      }
    });
  });
}
function getTime() {
  let today = new Date();
  const hours = today.getHours() >= 10 ? today.getHours() : "0" + today.getHours();
  const minutes = today.getMinutes() >= 10 ? today.getMinutes() : "0" + today.getMinutes();
  const seconds = today.getSeconds() >= 10 ? today.getSeconds() : "0" + today.getSeconds();
  return hours + "-" + minutes + "-" + seconds;
}
async function writeToExcel() {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet("My Sheet");

  worksheet.columns = [
    { header: "Team Name", key: "name", width: 10 },
    { header: "Score", key: "score", width: 32 },
  ];
  teams.map((team) => {
    worksheet.addRow({ name: team.name, score: team.score });
  });
  let today = new Date();
  let date = today.getFullYear() + "-" + (today.getMonth() + 1) + "-" + today.getDate();
  let fileName = "results/results - " + date + " " + getTime() + ".xlsx";
  await workbook.xlsx.writeFile(fileName);
}
async function insertData() {
  const serviceAccount = require("./key.json");

  fs.initializeApp({
    credential: fs.credential.cert(serviceAccount),
  });

  const db = fs.firestore();
  teams.map(async (team) => {
    const data = {
      teamName: team.name,
      score: team.score,
    };
    await db
      .collection("scores")
      .doc(data.teamName)
      .set(data)
      .then((res) => {})
      .catch((err) => {
        console.log(err);
      });
  });
}
async function main() {
  await getTotal();
  await writeToExcel();
  console.log("Excel Sheet Generated");
  await insertData();
  console.log("Data Inserted");
}
main().then();
