const cors = require("cors")({origin: true});
const { QueryTypes } = require("sequelize");
const excel = require('excel4node');

const {getGradesDataNoAsync, convertMarksToGrade} = require('../../utils/grade');
const {convertToMarathiNumber} = require('../../utils/translate');
const sequelize = require("../../utils/dbConnection");
const User = require("../../models/user");
const Class = require("../../models/class");
const School = require("../../models/school");
const Subject = require("../../models/subject");
const Student = require("../../models/student");
const Division = require("../../models/division");
const ClassTeacher  = require("../../models/classTeacher");
const EvalParam = require("../../models/evalParam");
const AcadamicYear = require("../../models/acadamicYear");
const DivisionSubject = require("../../models/divisionSubject");
const Assessment = require("../../models/assessment");
const AssessmentSubject = require("../../models/assessmentSubject");
const AssessmentSubjectEvalParam = require("../../models/assessmentSubjectEvalParam");
const AssessmentSubjectEvalResult = require("../../models/assessmentSubjectEvalResult");
//contants
const assessmentEvalParamKey = "aep";
const reportCardMarathiUnicode = "गुणपत्रक";
const subjectMarathiUnicode = "विषय";
const vSignUnicode = "व. सही"
const phSignUnicode = "प. सही"
const muSignUnicode = "मु. सही"
const totalMarathiUnicode = "एकूण गुण";
const percentageMarathiUnicode = "शे. गुण";
const gradeMarathiUnicode = "श्रेणी";
const serialNoMarathiUnicode = "क्र";
const studentNameMarathiUnicode = "विद्यार्थ्याचे नाव";
const marksMarathiUnicode = "गुण";
const classTeacherMarathiUnicode = "वर्गशिक्षक";
const classMarathiUnicode = "इयत्ता";
const averageMarathiUnicode = "सरासरी";
const overallResultMarathiUnicode = "निकालपत्रक";
const marathiLanguageCode = 'mr-IN';
const ignoreAvgColumns = ['roll_no', 'student_name'];
const studentPrefix = 'std';
const grades = getGradesDataNoAsync();

//Helper Variables 
let studentAverage = {};
let groupSubjectStudents = [];
//Helper Functions
const addStudentColumnData = (worksheet, key, columnIdx, startingRow, students, style, cache = false, groupSubjectCache = false) => {  
  let total = 0;
  let totalCount = 0;
  let includeAvg = false;
  if(ignoreAvgColumns.indexOf(key) == -1){
    includeAvg = true;
  }
  let i = 0;
  for(let student of students){
    let value = student[key];
    let data = student[key];
    let isNumber = false;


    if(typeof value == "number"){
      isNumber = true;
      data = convertToMarathiNumber(value + '');
    }else if(key == 'grade'){
      isNumber = true;
      value = student['percentage'] || 0;
    }

    worksheet.cell(startingRow, columnIdx)
      .string(data || '-')
      .style(style);

    if(includeAvg && isNumber){
      total += value;
      totalCount += 1;

      //We will add cache as true for overall report generation for total and percentage
      if(cache){
        const stdKey = `${studentPrefix}-${student.id}`;
        if(studentAverage[stdKey] == undefined){
          studentAverage[stdKey] = {
            percentage:0,
            total:0,
            percentageCount:0,
            totalCount:0
          };
          studentAverage[stdKey][key] = value;
          studentAverage[stdKey][key+'Count'] = 1;          
        }else{
          studentAverage[stdKey][key] += value;
          studentAverage[stdKey][key+'Count'] += 1;
        }
      }
    }

    if(groupSubjectCache){
      if(key == 'total'){
        if(groupSubjectStudents[i] == undefined){
          groupSubjectStudents[i] = {total:value};
        }else{
          groupSubjectStudents[i].total += value;
        }
      }
    }

    startingRow++;
    i++;
  }
  if(includeAvg && totalCount > 0){
    const average = Number(total / totalCount).toFixed(2);
    if(key == 'grade'){
      worksheet.cell(startingRow, columnIdx)
        .string(convertMarksToGrade(average, grades))
        .style(style);
    }else{
      worksheet.cell(startingRow, columnIdx)
        .string(convertToMarathiNumber(average))
        .style(style);
    }    
  }
  if(key == 'student_name'){
    worksheet.cell(startingRow, columnIdx)
      .string(averageMarathiUnicode)
      .style(style);
  }
  return includeAvg;
}

const addOverallStudentAverage = (worksheet, columnIdx, startingRow, students, style) => {  
  let totalAverage = 0;
  let totalAveragePercentage = 0;
  let averageGrade = '';

  let totalStudentCount = 0;

  let allStudentTotal = 0;
  let allStudentPercentage = 0;

  for(let student of students){
    const stdKey = `${studentPrefix}-${student.id}`;
    const data = studentAverage[stdKey];
    let percentage = 0;
    let percentageCount = 1;
    let total = 0;
    let totalCount = 1;
    let averageTotal = 0;
    let averagePercentage = 0;

    if(data != undefined){
      total = data.total;
      percentage = data.percentage;
      totalCount = data.totalCount == 0 ? 1 : data.totalCount;
      percentageCount = data.percentageCount == 0 ? 1 : data.percentageCount;

      averageTotal = total / totalCount;
      averagePercentage = percentage / percentageCount;

      allStudentTotal += averageTotal;
      allStudentPercentage += averagePercentage;

      totalStudentCount++;
    }
    const grade = convertMarksToGrade(averagePercentage, grades);

    worksheet.cell(startingRow, columnIdx)
      .string(convertToMarathiNumber(Number(total).toFixed(2)))
      .style(style);

    worksheet.cell(startingRow, columnIdx + 1)
      .string(convertToMarathiNumber(Number(averagePercentage).toFixed(2)))
      .style(style);

    worksheet.cell(startingRow, columnIdx + 2)
      .string(grade)
      .style(style);

    startingRow++;
  }
  if(totalStudentCount > 0){
       totalAverage = Number(allStudentTotal / totalStudentCount).toFixed(2);
       totalAveragePercentage = Number(allStudentPercentage / totalStudentCount).toFixed(2);
       averageGrade = convertMarksToGrade(totalAveragePercentage, grades);
  }else{
    totalAverage = '-';
    totalAveragePercentage = '-';
    averageGrade = '-'
  }
  worksheet.cell(startingRow, columnIdx)
    .string(convertToMarathiNumber(totalAverage))
    .style(style);

  worksheet.cell(startingRow, columnIdx + 1)
    .string(convertToMarathiNumber(totalAveragePercentage))
    .style(style);

  worksheet.cell(startingRow, columnIdx + 2)
    .string(averageGrade)
    .style(style);
}

const formatToExcel = async (data, callback) => { 
  const workbook = new excel.Workbook();
  const subjects = data?.subjects;
  if(subjects == undefined){
    console.log('No subjects Found!');
    return false;
  }
  const style = workbook.createStyle({
    font: {
      color: '#000000',
      size: 12
    }
  });
  const centerStyle = workbook.createStyle({
    font: {
      color: '#000000',
      size: 12
    },
    alignment: {
      wrapText: true,
      horizontal: 'center',
    }
  });

  //---------------Individual Subject Sheet---------------
  for(let subject of subjects){
    var worksheetRow = 1;
    if(subject?.columns == undefined){
      console.log('Subject Skipped As No Evaluation Paramaters Found!');
      return false;
    }
    const worksheet = workbook.addWorksheet(subject.name, {
      sheetFormat : {baseColWidth: 15,defaultRowHeight:20}
    });

    //School Title
    worksheet.cell(worksheetRow, 1, worksheetRow, subject.columns.length, true)
      .string(data.schoolTitle)
      .style(centerStyle);

    //Report Card
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, subject.columns.length, true)
      .string(reportCardMarathiUnicode)
      .style(centerStyle);

    //Exam Name & Academic Year
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, subject.columns.length, true)
      .string(`${subject.examName} - ${data.acadamicYear}`)
      .style(centerStyle);

    //Subject Name
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, subject.columns.length, true)
      .string(`${subjectMarathiUnicode} - ${subject.name}`)
      .style(centerStyle);

    //Class
    const halfColumns = parseInt(subject.columns.length / 2)
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, halfColumns, true)
      .string(`${classMarathiUnicode} - ${data.class}`)
      .style(centerStyle);

    //Class Teacher Name
    worksheet.cell(worksheetRow, halfColumns + 1, worksheetRow, subject.columns.length, true)
      .string(`${classTeacherMarathiUnicode} - ${data.classTeacherName}`)
      .style(centerStyle);

    //Columns Like Sr.No, Student Name....Total
    worksheetRow++;
    for (var i = 1; i <= subject.columns.length; i++) {
      const idx = i - 1;
      worksheet.cell(worksheetRow, i)
        .string(subject.columns[idx].name)
        .style(style);
      //Student Data Row like Serial Number...Marks...Total
      addStudentColumnData(
        worksheet, //Worksheet
        subject.columns[idx].key, //Column Key
        i, //Excel Column Index
        worksheetRow + 1, //Row Index
        subject.studentColumnWiseData, //Student List
        style //Cell Data Style
      );
    }

    worksheetRow += subject.studentColumnWiseData.length; //Setting Row After Student Rows
    worksheetRow += 3;//Adding Gap

    //Setting Footer Data Like V. Sign
    worksheet.cell(worksheetRow, 2)
      .string(vSignUnicode)
      .style(style);

    worksheet.cell(worksheetRow, 3)
      .string(phSignUnicode)
      .style(style);

    worksheet.cell(worksheetRow, subject.columns.length)
      .string(muSignUnicode)
      .style(style);

    //As We Know Second All is Student Name 
    //So we are increasing column width
    worksheet.column(2).setWidth(30);
  }
  //---------------Individual Subject Sheet---------------

  //---------------Group Subjects Sheet---------------
  const ignoreAlreadyCreatedSubjects = [];
  const groupSubjects = [];
  for(let subject of data.subjects){
    const subjectGroup = data.subjects.filter(s => s.group == subject.group);
    if(subjectGroup.length > 1){
      const groupName = subjectGroup[0].group;
      if(ignoreAlreadyCreatedSubjects.indexOf(groupName) == -1){
        ignoreAlreadyCreatedSubjects.push(groupName);
        const finalGroup = [];
        let i = 0;
        for(let subject of subjectGroup){
          let newColumns = [];
          if(i == 0){
            newColumns = subject.columns.filter(({key}) => key != 'percentage' && key != 'grade');
          }else{
            newColumns = subject.columns.filter(({key}) => key != 'percentage' && key != 'grade' && key != 'roll_no' && key != 'student_name');
          }
          newColumns.push({
            name : averageMarathiUnicode,
            key : 'average'
          })
          subject.columns = newColumns;
          finalGroup.push(subject);
          i++;
        }
        groupSubjects.push(finalGroup);        
      }      
    }
  }
  for(let subjects of groupSubjects){
    var worksheetRow = 1;
    var totalColumns = 0;
    const groupName = subjects[0].groupTranslated;
    const examName = subjects[0].examName;
    const worksheet = workbook.addWorksheet(groupName, {
      sheetFormat : {baseColWidth: 15,defaultRowHeight:20}
    });
    for(let subject of subjects){
      totalColumns += subject.columns.length;
    }
    totalColumns += 3 //Adding additional columns like Total, Percentage, Grade

    //School Title
    worksheet.cell(worksheetRow, 1, worksheetRow, totalColumns, true)
      .string(data.schoolTitle)
      .style(centerStyle);

    //Report Card
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, totalColumns, true)
      .string(reportCardMarathiUnicode)
      .style(centerStyle);

    //Exam Name & Academic Year
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, totalColumns, true)
      .string(`${examName} - ${data.acadamicYear}`)
      .style(centerStyle);

    //Subject Name
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, totalColumns, true)
      .string(`${subjectMarathiUnicode} - ${groupName}`)
      .style(centerStyle);

    //Class
    const halfColumns = parseInt(totalColumns / 2)
    worksheetRow++;
    worksheet.cell(worksheetRow, 1, worksheetRow, halfColumns, true)
      .string(`${classMarathiUnicode} - ${data.class}`)
      .style(centerStyle);

    //Class Teacher Name
    worksheet.cell(worksheetRow, halfColumns + 1, worksheetRow, totalColumns, true)
      .string(`${classTeacherMarathiUnicode} - ${data.classTeacherName}`)
      .style(centerStyle);

    //Columns Like Sr.No, Student Name....Total
    worksheetRow++;
    let columnIndex = 1;
    let maxMarks = 1;
    for(let subject of subjects){
      maxMarks = subject?.totalMaxMarks || maxMarks;
      for (var i = 1; i <= subject.columns.length; i++) {
        const idx = i - 1;
        worksheet.cell(worksheetRow, columnIndex)
          .string(subject.columns[idx].name)
          .style(style);
        //Student Data Row like Serial Number...Marks...Total
        addStudentColumnData(
          worksheet, //Worksheet
          subject.columns[idx].key, //Column Key
          columnIndex, //Excel Column Index
          worksheetRow + 1, //Row Index
          subject.studentColumnWiseData, //Student List
          style, //Cell Data Style,
          false,
          true
        );
        columnIndex += 1;
      }
    }
    maxMarks = maxMarks * subjects.length;

    //Percentage & Grade Calculation    
    let finalStudent = [];
    for(let gsStudent of groupSubjectStudents){
      const percentage = (gsStudent.total / maxMarks) * 100;
      gsStudent.percentage = parseFloat(percentage.toFixed(2));
      gsStudent.grade = convertMarksToGrade(percentage, grades);
      finalStudent.push(gsStudent);
    }

    //Total Data For Group Subject
    worksheet.cell(worksheetRow, columnIndex)
      .string(totalMarathiUnicode)
      .style(style);    
    addStudentColumnData(
      worksheet, //Worksheet
      "total", //Column Key
      columnIndex, //Excel Column Index
      worksheetRow + 1, //Row Index
      finalStudent, //Student List
      style, //Cell Data Style,
      false,
      false
    );

    //Percentage For Group Subject
    columnIndex++;
    worksheet.cell(worksheetRow, columnIndex)
      .string(percentageMarathiUnicode)
      .style(style);    
    addStudentColumnData(
      worksheet, //Worksheet
      "percentage", //Column Key
      columnIndex, //Excel Column Index
      worksheetRow + 1, //Row Index
      finalStudent, //Student List
      style, //Cell Data Style,
      false,
      false
    );

    //Grade For Group Subject
    columnIndex++;
    worksheet.cell(worksheetRow, columnIndex)
      .string(gradeMarathiUnicode)
      .style(style);    
    addStudentColumnData(
      worksheet, //Worksheet
      "grade", //Column Key
      columnIndex, //Excel Column Index
      worksheetRow + 1, //Row Index
      finalStudent, //Student List
      style, //Cell Data Style,
      false,
      false
    );

    worksheetRow += subjects[0].studentColumnWiseData.length; //Setting Row After Student Rows
    worksheetRow += 3;//Adding Gap

    //Setting Footer Data Like V. Sign
    worksheet.cell(worksheetRow, 2)
      .string(vSignUnicode)
      .style(style);

    worksheet.cell(worksheetRow, 3)
      .string(phSignUnicode)
      .style(style);

    worksheet.cell(worksheetRow, totalColumns)
      .string(muSignUnicode)
      .style(style);    

    //As We Know Second All is Student Name 
    //So we are increasing column width
    worksheet.column(2).setWidth(30);
    groupSubjectStudents = [];
  }
  //---------------Group Subjects Sheet---------------


  //---------------Overall Subject Sheet---------------
  //Overall Result
  var worksheetRow = 1;
  var worksheetCol = 1;
  const totalColumnsForOverall = (data.subjects.length * 3) + (2 + 3);
  //2 For RollNo and Student name
  //3 For Total Marks, Percentage, Grade
  const worksheet = workbook.addWorksheet(overallResultMarathiUnicode, {
      sheetFormat : {baseColWidth: 15,defaultRowHeight:20}
  });

  //School Title
  worksheet.cell(worksheetRow, 1, worksheetRow, totalColumnsForOverall, true)
    .string(data.schoolTitle)
    .style(centerStyle);

  //Exam Name & Academic Year
  worksheetRow++;
  worksheet.cell(worksheetRow, 1, worksheetRow, totalColumnsForOverall, true)
    .string(`${data.acadamicYear}`)
    .style(centerStyle);

  //Overall Result Text
  worksheetRow++;
  worksheet.cell(worksheetRow, 1, worksheetRow, totalColumnsForOverall, true)
    .string(overallResultMarathiUnicode)
    .style(centerStyle);

  //Class
  const halfColumns = parseInt(totalColumnsForOverall / 2)
  worksheetRow++;
  worksheet.cell(worksheetRow, 1, worksheetRow, halfColumns, true)
    .string(`${classMarathiUnicode} - ${data.class}`)
    .style(centerStyle);

  //Class Teacher Name
  worksheet.cell(worksheetRow, halfColumns + 1, worksheetRow, totalColumnsForOverall, true)
    .string(`${classTeacherMarathiUnicode} - ${data.classTeacherName}`)
    .style(centerStyle);

  //Overall Header like Serial No....Total
  worksheetRow++;
  worksheet.cell(worksheetRow, worksheetCol, worksheetRow + 1, worksheetCol, true)
    .string(`${serialNoMarathiUnicode}`)
    .style(centerStyle);
  addStudentColumnData(
    worksheet, //Worksheet
    "roll_no", //Column Key
    worksheetCol, //Excel Column Index
    worksheetRow + 2, //Row Index
    data.subjects[0].studentColumnWiseData, //Student List
    style //Cell Data Style
  );
  worksheetCol++;
  worksheet.cell(worksheetRow, worksheetCol, worksheetRow + 1, worksheetCol, true)
    .string(`${studentNameMarathiUnicode}`)
    .style(centerStyle);
  addStudentColumnData(
    worksheet, //Worksheet
    "student_name", //Column Key
    worksheetCol, //Excel Column Index
    worksheetRow + 2, //Row Index
    data.subjects[0].studentColumnWiseData, //Student List
    style //Cell Data Style
  );

  for(let subject of data.subjects){
    worksheetCol++;
    worksheet.cell(worksheetRow, worksheetCol, worksheetRow, worksheetCol + 2, true)
      .string(subject.name)
      .style(centerStyle);

    worksheet.cell(worksheetRow + 1, worksheetCol)
      .string(totalMarathiUnicode)
      .style(centerStyle);

    addStudentColumnData(
      worksheet, //Worksheet
      "total", //Column Key
      worksheetCol, //Excel Column Index
      worksheetRow + 2, //Row Index
      subject.studentColumnWiseData, //Student List
      style,//Cell Data Style
      true
    );

    worksheet.cell(worksheetRow + 1, worksheetCol + 1)
      .string(percentageMarathiUnicode)
      .style(centerStyle);
    addStudentColumnData(
      worksheet, //Worksheet
      "percentage", //Column Key
      worksheetCol + 1, //Excel Column Index
      worksheetRow + 2, //Row Index
      subject.studentColumnWiseData, //Student List
      style,//Cell Data Style
      true
    );

    worksheet.cell(worksheetRow + 1, worksheetCol + 2)
      .string(gradeMarathiUnicode)
      .style(centerStyle);
    addStudentColumnData(
      worksheet, //Worksheet
      "grade", //Column Key
      worksheetCol + 2, //Excel Column Index
      worksheetRow + 2, //Row Index
      subject.studentColumnWiseData, //Student List
      style //Cell Data Style
    );

    worksheetCol += 2;
  }

  //Overall Average
  worksheetCol++;

  worksheet.cell(worksheetRow, worksheetCol, worksheetRow + 1, worksheetCol, true)
      .string(totalMarathiUnicode)
      .style(centerStyle);

  worksheet.cell(worksheetRow, worksheetCol + 1, worksheetRow + 1, worksheetCol + 1, true)
        .string(percentageMarathiUnicode)
        .style(centerStyle);

  worksheet.cell(worksheetRow, worksheetCol + 2, worksheetRow + 1, worksheetCol + 2, true)
        .string(gradeMarathiUnicode)
        .style(centerStyle);

  addOverallStudentAverage(worksheet, //Worksheet
    worksheetCol, //Excel Column Index
    worksheetRow + 2, //Row Index
    data.subjects[0].studentColumnWiseData, //Student List
    style //Cell Data Style
  );
  
  //Setting Footer Data Like V. Sign
  worksheetRow += data.subjects[0].studentColumnWiseData.length; //Setting Row After Student Rows
  worksheetRow += 4;//Adding Gap
  worksheet.cell(worksheetRow, 2)
    .string(vSignUnicode)
    .style(style);

  worksheet.cell(worksheetRow, 3)
    .string(phSignUnicode)
    .style(style);

  worksheet.cell(worksheetRow, totalColumnsForOverall)
    .string(muSignUnicode)
    .style(style);

  //---------------Overall Subject Sheet---------------

  workbook.writeToBuffer().then(callback);
}

const getReportData = async (divisionId, assessmentId, translateKey) => {    
  try {
    //Variable for getting already fetched evalParam
    let evalParamCache = {};

    //Division...
    const division = await Division.findOne({
      where: {
        id: divisionId
      }
    });
    if (!division) return "Division not found!";

    //Division...
    const divisionClass = await Class.findOne({
      attributes : ['className'],
      where: {
        id: division.classId
      }
    });
    if (!divisionClass) return "Class not found!";

    let teacher = {name : ''}
    const classTeacher = await ClassTeacher.findOne({
      attributes : ['teacherId'],
      where : {divisionId: divisionId}
    });    
    if(classTeacher){        
      const user = await User.findOne({
        attributes : ['firstName', 'middleName','lastName'],
        where : {authId : classTeacher.teacherId}
      });
      if(user){
        teacher.name = `${user.firstName} ${user.middleName} ${user.lastName}`;
      }
    }

    //School...
    const school = await School.findOne({
      attributes : ['id', 'title'],
      where: {
        id: division.schoolId
      }
    });
    if (!school) return "School not found!";

    //Assessment...
    const assessment = await Assessment.findOne({
      where: {
        id: assessmentId
      }
    });
    if (!assessment) return "Assessment not found!";

    //Academic Year...
    const acadamicYear = await AcadamicYear.findOne({
      attributes : ['title'],
      where: {
        id: division.acadamicYearId,        
      }
    });
    if (!acadamicYear) return "Academic year not found!";

    const assessmentSubjects = await AssessmentSubject.findAll({
      where : {
        assessmentId: assessmentId,
        divisionId : divisionId,          
      }
    });
    if(!assessmentSubjects || assessmentSubjects?.length == 0){
      return res.status(400).send({
        "message": "No Subject found for this exam!"
      });
    }

    const students = await Student.findAll({
      attributes : ['id', 'rollNo', 'firstName', 'middleName', 'lastName', 'additionalInfo'],
      order: [         
        ['roll_no', 'ASC'],
      ],
      where : {
        divisionId : divisionId,          
      }
    });
    if(!students || students?.length == 0){
      return res.status(400).send({
        "message": "No Students found in class!"
      });
    }

    var excelJSONData = {
      schoolTitle : school.title,
      acadamicYear : acadamicYear.title,
      class : convertToMarathiNumber(divisionClass.className),
      classTeacherName : teacher.name, 
    }
    var subjects = [];
    var overallSubjects = [];
    for(let assessmentSubject of assessmentSubjects){
      let maxMarks = 0;  

      const subjectData = await Subject.findOne({
        attributes : ['id', 'title', 'translations', 'group'],
        where : {id : assessmentSubject.subjectId}
      });      

      const params = await AssessmentSubjectEvalParam.findAll({
        attributes : ['id', 'evalParamId', 'maxMarks'],
        where : {assessmentSubjectId : assessmentSubject.id},
        order: [         
          ['relative_order', 'ASC'],
        ]
      });
      let subjectName = subjectData.title;
      let groupNameTranslated = subjectData.group;
      if(subjectData.translations != undefined){
        if(subjectData.translations?.subject != undefined && subjectData.translations.subject[translateKey] != undefined){
          subjectName = subjectData.translations.subject[translateKey];
        }
        if(subjectData.translations?.group != undefined && subjectData.translations.group[translateKey] != undefined){
          groupNameTranslated = subjectData.translations.group[translateKey];
        }
      }
      var subject = {
        name : subjectName,
        examName : assessment.title,
        group : subjectData.group,
        groupTranslated : groupNameTranslated,
        columns : [],
        studentColumnWiseData : []
      }
      const assessmentEvalParamIds = [];
      //Adding Serial No Column
      subject.columns.push({
          name : serialNoMarathiUnicode,
          key : 'roll_no'
      });
      //Adding Student Name Column
      subject.columns.push({
          name : studentNameMarathiUnicode,
          key : 'student_name'
      });
      //Setting Eval Params as Columns
      for(let param of params){
        if(param.maxMarks == null){
          //TODO: Default Param
          continue
        }
        const evalKey = `e${param.evalParamId}`;
        let evalParam = null;
        if(evalParamCache[evalKey] == undefined){
          evalParam = await EvalParam.findOne({
            where : {
              id : param.evalParamId,
            }           
          })
          evalParamCache[evalKey] = evalParam;
        }else{
          evalParam = evalParamCache[evalKey];
        }
        if(evalParam == undefined){
          //TODO: may be we should throw error
          continue
        }
        const maxMarksMarathi = convertToMarathiNumber((param.maxMarks || 0) + '');
        let title = `${evalParam.title} ${maxMarksMarathi} ${marksMarathiUnicode}`;
        if(translateKey != undefined && evalParam?.translations != undefined){
          if(evalParam.translations[translateKey] != undefined){
            title = `${evalParam.translations[translateKey]} ${maxMarksMarathi} ${marksMarathiUnicode}`;
          }
        }
        subject.columns.push({
          name : title,
          key : `${assessmentEvalParamKey}${param.id}`,
        });
        maxMarks += param.maxMarks || 0;
        assessmentEvalParamIds.push(param.id);
      }
      //Adding Total Column
      subject.columns.push({
          name : totalMarathiUnicode,
          key : 'total'
      });
      //Adding Percentage Column
      subject.columns.push({
          name : percentageMarathiUnicode,
          key : 'percentage'
      });
      //Adding Grade Column
      subject.columns.push({
          name : gradeMarathiUnicode,
          key : 'grade'
      });

      //Get Results for eval params
      for(let student of students){
        let total = 0;
        let average = 0;
        const rollNo = student.rollNo == null ? '-' : convertToMarathiNumber(`${student.rollNo}`);

        let studentName = `${student.firstName} ${student.middleName} ${student.lastName}`;
        if(translateKey == "marathi"){
          if(student?.additionalInfo?.translation != undefined){
            const translation = student?.additionalInfo?.translation;
            if(translation[marathiLanguageCode] != undefined){
              studentName = translation[marathiLanguageCode].studentName;
            }
          }
        }
        const studentMarkData = {
          id : student.id,
          roll_no: rollNo,
          student_name : studentName,
        };
        const resultParams = await AssessmentSubjectEvalResult.findAll({
          attributes : ['marks','assessmentSubjectEvalParamId'],
          where : {assessmentSubjectEvalParamId : assessmentEvalParamIds, studentId:student.id},
        });
        for(let resultParam of resultParams){
          let marks = resultParam.marks || 0;
          total += marks;
          studentMarkData[`${assessmentEvalParamKey}${resultParam.assessmentSubjectEvalParamId}`] = marks;
        }
        const percentage = total == 0 ? 0 : Math.round((parseInt(total) / parseInt(maxMarks)) * 100);
        const grade = convertMarksToGrade(percentage, grades);
        const resultParamsLength = resultParams.length == 0 ? 1 : resultParams.length;
        studentMarkData.total = total;
        studentMarkData.average = parseFloat((total / resultParamsLength).toFixed(2));
        studentMarkData.percentage = percentage;
        studentMarkData.grade = grade || '';
        subject.totalMaxMarks = maxMarks;
        subject.studentColumnWiseData.push(studentMarkData);        
      }

      subjects.push(subject);
    }

    excelJSONData['subjects'] = subjects;
    return excelJSONData;

  } catch (err) {
    console.log(err)
    return false;
  }
};
//Entry Point
const getResultReport = async (req, res) => {
  cors(req, res, async function () {
    try {
      const { divisionId, assessmentId, translateKey } = req.body;
      if(divisionId == undefined){
        res.status(400).send("Divsion id is required!");
      }else if(assessmentId == undefined){
        res.status(400).send("Assessment Id id is required!");
      }else if (translateKey == undefined) {
        res.status(400).send("Translate Key id is required!");
      }
      const reportData = await getReportData(divisionId, assessmentId, translateKey);
      if(reportData == false){
        res.status(500).send("Server Error!");
      }else if(typeof reportData == "string"){
        //If Response is string then there was an error
        res.status(400).send(reportData);
      }
      formatToExcel(reportData, (buffer) => {
        res.setHeader("Access-Control-Allow-Origin", "*");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=Exam_Report.xlsx");
        res.status(200).send(buffer);
      });
    } catch(err){
      console.log(err);
      res.status(500).send("Error In Generating File");
    }
  })
}

module.exports = getResultReport;

/*
  Sample Data Created From GetReportData Function

  {
    schoolTitle : 'Vinaj Parivar Vidhya Mandir', 
    acadamicYear : '2022-23',
    subjects : [
      {
        name : 'Maths',
        examName : 'Summative Assesment',        
        classTeacherName : 'Deepak Rathod',
        columns : [
          {
            name : 'Kra',
            key:'roll_no'
          },
          {
            name : 'Student Name',
            key:'student_name'
          },
          {
            name : 'Written Exam',
            key:'ET_1',
          },
          {
            name : 'Total',
            key:'total',
          },
          {
            name : 'Grade',
            key:'grade',
          },
        ],
        studentColumnWiseData : [
          {
            roll_no : "1",
            student_name : 'Deepak Rathod',
            ET_1:'90',
            total:"90",
            grade : 'B1'
          }
        ]
      }
    ]
  }
*/
