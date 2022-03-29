/* CREATE data in excel sheet and then GET data from that excel sheet */ 

const res = require("express/lib/response");
const http = require("http");
const xlsx = require("xlsx");
const PORT = 8000;

const server = http.createServer((req, res) => {
  if ((req.url === "/post-data", req.method === "POST")) {
    res.writeHead(200, { "Content-Type": "html/plain" });
    const students = [
      {
        StudentName: "Bharat",
        Course: "B.Tech(CSE)",
        email: "bharat@gmail.com",
        age: 23,
        CollegeName: "Vidya College Meerut",
      },
      {
        StudentName: "Sachin",
        Course: "B.Tech(CSE)",
        email: "sachin@gmail.com",
        age: 25,
        CollegeName: "MIET Meerut",
      },
      {
        StudentName: "Rohit Sharma",
        Course: "B.Tech(CSE)",
        email: "rohitsharma@gmail.com",
        age: 25,
        CollegeName: "SRGC Muzaffarnagar",
      },
    ];
    function dataSheet() {
      const workSheet = xlsx.utils.json_to_sheet(students);
      const workBook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(workBook, workSheet, "students");
      xlsx.writeFile(workBook, "studentsData.xlsx");
    }
    res.write("File Write Successfully!");
    res.end(dataSheet());
  } else if (req.url == "/get-data" && req.method === "GET") {
    res.writeHead(200, { "Content-Type": "html/JSON" });
    var workbook = xlsx.readFile(__dirname + "/studentsData.xlsx");
    var sheet_name_list = workbook.SheetNames;
    var xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    res.end(JSON.stringify(xlData));
    console.log("Sheet Open Successfully in Postman !");
  }
});

server.listen(PORT, () => {
  console.log(`Server run on ${PORT}`);
});
