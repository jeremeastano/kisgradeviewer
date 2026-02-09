/**
 * REQUIRED ENTRY POINT FOR WEB APP
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index.html')
    .setTitle('KIS Online Grade Viewer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * FETCH STUDENT DATA
 */
function submitData(obj) {
  var spreadsheetId = "1Z3bNZ8aHX1Ll-qhI2lQe1zKqt-eQYCPdHTslBZ5-PIw";
  var sheetName = "Grades";

  if (!obj) {
    return "<div style='color:red;'>Invalid access code.</div>";
  }

  var accessCode = String(obj).trim(); 
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === accessCode) {

      var name = data[i][1] || "-";
      var lrn = data[i][2] || "N/A";
      var sectionInfo = data[i][4];
      var schoolyearInfo = data[i][3];
      var generalAve = data[i][45];
      var remarks = data[i][46] || "—";

      var subjects = [
        { name: "FILIPINO", q1: data[i][5], q2: data[i][6], q3: data[i][7], q4: data[i][8], ave: data[i][9] },
        { name: "ENGLISH",  q1: data[i][10], q2: data[i][11], q3: data[i][12], q4: data[i][13], ave: data[i][14] },
        { name: "MATH",     q1: data[i][15], q2: data[i][16], q3: data[i][17], q4: data[i][18], ave: data[i][19] },
        { name: "SCIENCE",  q1: data[i][20], q2: data[i][21], q3: data[i][22], q4: data[i][23], ave: data[i][24] },
        { name: "AP",       q1: data[i][25], q2: data[i][26], q3: data[i][27], q4: data[i][28], ave: data[i][29] },
        { name: "EsP",      q1: data[i][30], q2: data[i][31], q3: data[i][32], q4: data[i][33], ave: data[i][34] },
        { name: "TLE",      q1: data[i][35], q2: data[i][36], q3: data[i][37], q4: data[i][38], ave: data[i][39] },
        { name: "MAPEH",    q1: data[i][40], q2: data[i][41], q3: data[i][42], q4: data[i][43], ave: data[i][44] }
      ];

      // --- CALCULATE QUARTERLY AVERAGES & N/A CHECK ---
      var qStats = [
        { sum: 0, valid: true, key: 'q1', color: '#3498db' }, // Blue
        { sum: 0, valid: true, key: 'q2', color: '#2ecc71' }, // Green
        { sum: 0, valid: true, key: 'q3', color: '#f1c40f' }, // Yellow
        { sum: 0, valid: true, key: 'q4', color: '#e74c3c' }  // Red
      ];

      subjects.forEach(function(s) {
        qStats.forEach(function(q) {
          var val = s[q.key];
          // Check if value is empty, null, or not a valid number
          if (val === "" || val === null || isNaN(parseFloat(val))) {
            q.valid = false;
          } else {
            q.sum += parseFloat(val);
          }
        });
      });

      // --- START HTML OUTPUT ---
      var html = '<p style="margin: 25px 0 15px 0; font-size: 17px; color: #354a5f; font-weight: bold; text-align: center;">Student Academic Progress Report</p>';

      html += '<div class="report-header">' +
                 '<div><strong>NAME:</strong> ' + name + '</div>' +
                 '<div><strong>LRN:</strong> ' + lrn + '</div>' +
                 '<div><strong>SCHOOL YEAR:</strong> ' + schoolyearInfo + '</div>' +
                 '<div><strong>GRADE & SECTION:</strong> ' + sectionInfo + '</div>' +
                 '</div>';

      html += '<table><thead><tr><th>SUBJECTS</th><th>Q1</th><th>Q2</th><th>Q3</th><th>Q4</th><th>AVE</th></tr></thead><tbody>';

      subjects.forEach(function(sub) {
        var displayAve = (typeof sub.ave === 'number') ? sub.ave.toFixed(0) : (sub.ave || '-');
        html += '<tr><td class="sub-name">' + sub.name + '</td>' +
                '<td>' + (sub.q1 || '-') + '</td><td>' + (sub.q2 || '-') + '</td>' +
                '<td>' + (sub.q3 || '-') + '</td><td>' + (sub.q4 || '-') + '</td>' +
                '<td><strong>' + displayAve + '</strong></td></tr>';
      });

      var roundedGeneralAve = (typeof generalAve === 'number') ? generalAve.toFixed(0) : (generalAve || "N/A");
      html += '<tr class="total-row"><td>GENERAL AVERAGE</td><td colspan="4"></td><td>' + roundedGeneralAve + '</td></tr>';
      html += '</tbody></table>';

      // --- GENERATE QUARTERLY BLOCKS (ORIGINAL COLORS MAINTAINED) ---
      html += '<div style="display: flex; gap: 10px; margin-top: 20px; flex-wrap: wrap; font-family: sans-serif;">';
      
      qStats.forEach(function(q, index) {
        var displayVal = q.valid ? (q.sum / subjects.length).toFixed(0) : "N/A";
        
        html += '<div style="flex: 1; min-width: 100px; background: ' + q.color + '; color: white; padding: 15px; border-radius: 8px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' +
          '<div style="font-size: 11px; opacity: 0.9; font-weight: bold;">AVERAGE Q' + (index + 1) + '</div>' +
          '<div style="font-size: 22px; font-weight: bold;">' + displayVal + '</div>' +
        '</div>';

      });
      html += '</div>';

      html += '<div class="remarks-box"><strong>REMARKS:</strong> ' + remarks + '</div>';

      var aiInterpretation = generateAIInterpretation(generalAve, subjects);
      html += '<div style="margin: 20px 0; padding: 20px; background-color: #fcfcfc; border-left: 8px solid #add8e6; font-size: 14px; line-height: 1.8; color: #333; text-align: justify; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">' + aiInterpretation + '</div>';

      html += '<div style="margin: 20px auto; border: 2px dashed #ccc; padding: 10px; max-width: 500px; display: table;">';
      html += '<div style="text-align: center; font-weight: bold; font-size: 13px; margin-bottom: 8px;">GRADING SCALE TABLE / RUBRIC</div>';
      html += '<table style="width: 100%; border-collapse: collapse; font-size: 12px; text-align: center;">';
      html += '<thead><tr>' +
                '<th style="padding: 5px; text-align: left; background-color: #f2f2f2; border: 1px solid #ddd; font-weight: bold; color: #000000;">Descriptors</th>' +
                '<th style="padding: 5px; background-color: #f2f2f2; border: 1px solid #ddd; font-weight: bold; color: #000000;">Grading Scale</th>' +
                '<th style="padding: 5px; background-color: #f2f2f2; border: 1px solid #ddd; font-weight: bold; color: #000000;">Remarks</th>' +
              '</tr></thead>';
      html += '<tbody>' +
                '<tr><td style="padding: 5px; text-align: left; border: 1px solid #ddd;">Outstanding</td><td style="border: 1px solid #ddd;">90-100</td><td style="border: 1px solid #ddd;">Passed</td></tr>' +
                '<tr><td style="padding: 5px; text-align: left; border: 1px solid #ddd;">Very Satisfactory</td><td style="border: 1px solid #ddd;">85-89</td><td style="border: 1px solid #ddd;">Passed</td></tr>' +
                '<tr><td style="padding: 5px; text-align: left; border: 1px solid #ddd;">Satisfactory</td><td style="border: 1px solid #ddd;">80-84</td><td style="border: 1px solid #ddd;">Passed</td></tr>' +
                '<tr><td style="padding: 5px; text-align: left; border: 1px solid #ddd;">Fairly Satisfactory</td><td style="border: 1px solid #ddd;">75-79</td><td style="border: 1px solid #ddd;">Passed</td></tr>' +
                '<tr><td style="padding: 5px; text-align: left; border: 1px solid #ddd;">Did Not Meet Expectations</td><td style="border: 1px solid #ddd;">Below 75</td><td style="border: 1px solid #ddd;">Failed</td></tr>' +
              '</tbody></table></div>';

      html += '<div style="margin-top:20px; text-align:center; font-size:14px; color:red; line-height:1.4;">' +
              'Note: The copy of grades from this portal is not a substitute for the Report Card (School Form 9) of the Department of Education. This document should not be used for any school transactions.' +
              '</div>';

      html += '<div style="margin-top:20px; text-align:center;">' +
              '<button onclick="window.print()" style="padding:10px 20px; font-size:16px; cursor:pointer; background-color: #354a5f; color: white; border: none; border-radius: 4px;">Print / Save as PDF</button>' +
              '</div>';

      return html;
    }
  }
  return "<div style='color:red; margin-top:20px; text-align: center;'>Student record not found. Please check your access code and try again.</div>";
}

/**
 * HELPER FOR DESCRIPTOR LOGIC
 */
function getDescriptor(avg) {
  if (avg >= 90) return "Outstanding";
  if (avg >= 85) return "Very Satisfactory";
  if (avg >= 80) return "Satisfactory";
  if (avg >= 75) return "Fairly Satisfactory";
  return "Did Not Meet Expectations";
}

/**
 * TREND-BASED INTERPRETATION ENGINE
 */
function generateAIInterpretation(avg, subjects) {
  var activeSubjects = subjects.filter(function(s) { 
    return (typeof s.q1 === 'number' || typeof s.q2 === 'number' || typeof s.q3 === 'number' || typeof s.q4 === 'number'); 
  });
  
  if (activeSubjects.length === 0) return "<div style='text-align:center; padding: 20px; color: #666;'>Analyzing trends... Please wait for quarterly grades to be encoded.</div>";

  var improvingCount = 0;
  var decliningCount = 0;
  var displayAvg = typeof avg === 'number' ? avg : 0;
  var descriptor = getDescriptor(displayAvg);
  
  var cardStyle = "background: #ffffff; border: 1px solid #e1e8ed; border-radius: 10px; padding: 18px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);";
  var titleStyle = "color: #2c3e50; font-size: 14px; font-weight: 800; letter-spacing: 1px; margin-bottom: 12px; display: flex; align-items: center; text-transform: uppercase;";
  var iconCircle = "display: inline-block; width: 28px; height: 28px; background: #eef2f7; border-radius: 50%; text-align: center; line-height: 28px; margin-right: 10px; font-size: 16px;";

  var html = "";

  html += '<div style="' + cardStyle + '">';
  html += '<div style="' + titleStyle + '"><span style="' + iconCircle + '">📈</span> Overall Trend</div>';
  
  var statusColor = (displayAvg >= 90) ? "#27ae60" : (displayAvg >= 75) ? "#2980b9" : "#e74c3c";
  html += "<p style='margin: 0; color: #444;'>The student's performance is currently categorized as <strong style='color:" + statusColor + ";'>" + (displayAvg > 0 ? descriptor.toUpperCase() : "IN PROGRESS") + "</strong>. ";
  
  activeSubjects.forEach(function(sub) {
    var grades = [sub.q1, sub.q2, sub.q3, sub.q4].filter(function(g) { return typeof g === 'number' && g > 0; });
    if (grades.length >= 2) {
      var last = grades[grades.length - 1];
      var prev = grades[grades.length - 2];
      if (last > prev) improvingCount++;
      else if (last < prev) decliningCount++;
    }
  });

  if (improvingCount > decliningCount) html += "There is a notable <strong style='color:#27ae60;'>upward momentum</strong> in the student's academic journey.</p>";
  else if (decliningCount > improvingCount) html += "The student is facing some <strong style='color:#e67e22;'>challenges</strong> in maintaining previous grade levels.</p>";
  else html += "The student is maintaining a <strong style='color:#34495e;'>steady and consistent</strong> academic pace.</p>";
  html += '</div>';

  html += '<div style="' + cardStyle + '">';
  html += '<div style="' + titleStyle + '"><span style="' + iconCircle + '">📋</span> Subject-by-Subject Analysis</div>';
  html += '<div style="display: grid; gap: 10px;">';
  
  activeSubjects.forEach(function(sub) {
    var grades = [sub.q1, sub.q2, sub.q3, sub.q4].filter(function(g) { return typeof g === 'number' && g > 0; });
    var currentSubAve = (typeof sub.ave === 'number') ? sub.ave : (grades.reduce((a, b) => a + b, 0) / (grades.length || 1));
    var subDescriptor = getDescriptor(currentSubAve);
    var subColor = currentSubAve >= 75 ? "#2c3e50" : "#c0392b";
    
    html += '<div style="padding: 10px 15px; background: #f8fbff; border-radius: 6px; border-left: 4px solid #3498db;">';
    html += '<span style="font-weight: bold; color: ' + subColor + ';">' + sub.name + '</span>: ';
    html += 'The student received a rating of <strong>' + subDescriptor +
        '</strong> and has <i>' + (currentSubAve >= 75 ? 'passed the subject' : 'not met the subject requirements') + '</i>. ';

    if (grades.length >= 2) {
      var last = grades[grades.length - 1];
      var prev = grades[grades.length - 2];
      if (last > prev) html += "<span style='color: #27ae60; font-size: 12px;'>↑ Improving</span>";
      else if (last < prev) html += "<span style='color: #e74c3c; font-size: 12px;'>↓ Slight Dip</span>";
    }
    html += '</div>';
  });
  html += '</div></div>';

  html += '<div style="' + cardStyle + ' border-left: 5px solid #f1c40f; background: #fffdf5;">';
  html += '<div style="' + titleStyle + '"><span style="' + iconCircle + '">💡</span> Teacher Suggestions</div>';
  html += '<div style="color: #5d4037; line-height: 1.6;">';
  
  if (decliningCount > 0) {
    html += "To reverse declining trends, focus on <strong>consistent improvement of study habits</strong>. Ensure adequate preparation for examinations, outline lessons to reinforce understanding, and manage time effectively to submit quality outputs aligned with the provided rubric.";
  } else if (improvingCount > 0 && decliningCount === 0) {
    html += "<strong>Keep up the excellent momentum!</strong> Maintaining this current study routine and continuing to submit quality work will help secure higher averages in the coming quarters.";
  } else if (displayAvg < 75 && displayAvg > 0) {
    html += "A more structured study plan and additional academic support are recommended to help bring grades up to a passing level.";
  } else {
    html += "<strong>Consistent practice and active participation in class are key.</strong> Focusing on steady progress across all subjects will help improve the overall academic standing.";
  }
  html += '</div></div>';

  return html;
}
