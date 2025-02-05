// --- Start: Remote Data Helpers --- //

function parseCSV(csvString, ignoreHeaders) {
  // --- simple csv parser modified from: modjeska.us/csv-google-sheets-basic-auth --- //
  var final = [];
  var lines = csvString.split(/\n/g);
  var start = ignoreHeaders ? 1 : 0;
  for (var i = start; i < lines.length; i++) {
    var line = lines[i];
    if ( line == '') {
      final.push([]);
    } else {
      final.push(line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/));
    }
  }
  return final;
}

function embedRemoteData(url, base64pw, randStr, csvParse, ignoreHeaders) {
  // --- embeds remote data (csv & otherwise) from remote url w options (authentication, refresh, csv parse) --- //
  // --- based on original script written by: modjeska.us/csv-google-sheets-basic-auth --- //
  let headers = {};
  if ( base64pw ) {
    headers['Authorization'] = 'Basic '.concat(base64pw);
  }
  var resp = UrlFetchApp.fetch(url, { headers: headers });
  if ( csvParse ) {
    var csvContent = parseCSV(resp.getContentText(), ignoreHeaders);
    return csvContent;
  }
  return resp.getContentText();
}

// --- End: Remote Data Helpers --- //


// === Workout Manager Main === //


// -- Helpers: alrt(), splitAtLast(), convertDate() -- //

function alrt(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

function splitAtLast(str, char) { /* via googai */
  const lastIndex = str.lastIndexOf(char);
  if ( lastIndex === -1 ) {
    return [str];
  }
  return [str.slice(0, lastIndex), str.slice(lastIndex + 1)];
}

function convertDate(dateString) { /* via googai */
  const date = new Date(dateString);
  const month = date.toLocaleString('default', { month: 'short' });
  const day = date.getDate();
  return `${month} ${day}`;
}

function currentTime(format = 'short') {
  // --- update time set: current date & time in format: 1/1 12:30:06p --- //
  const date = new Date()
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const year = date.getFullYear().toString().slice(-2);
  const hours = date.getHours();
  const minutes = date.getMinutes().toString().padStart(2, '0');
  const seconds = date.getSeconds().toString().padStart(2, '0');
  const period = hours >= 12 ? 'p' : 'a';
  const formattedHours = hours % 12 || 12;
  let output = '';
  if ( format == 'short' ) {
    output = `${month}/${day} ${formattedHours}:${minutes}${period}`;
  } else if ( format == 'long' ) {
    output = `${month}/${day}/${year} ${formattedHours}:${minutes}:${seconds}${period}`;
  }
  return output;
}

// -- End: Helpers -- //

// -- Sheet Cell Helpers: convertCellLocationTags(), formatCellsList, formatCellsAdd(), formatCellsApply(), renderCSVList() -- //

function convertCellLocationTags(cell, r, c) {
  // --- receives a cell value string and converts cell location tags (e.g. $cellref, $colref) to A1 references --- //
  var alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split(''); // only 26 columns supported
  // $cellref(r, c)
  cell = cell.replace(/\$cellref\(([\d-]+),([\d-]+)\)/g, (match, row, col) => {
    // r[0] = row 1
    // c[0] = col A
    ref_row = parseInt(row) + 1;
    ref_col = parseInt(col);
    return `${alpha[c+ref_col]}${r+ref_row}`;
  });
  // $colref(c) = A, B, C, ...
  cell = cell.replace(/\$colref\(([\d-]+)\)/g, (match, col) => {
    // c[0] = col A
    ref_col = parseInt(col);
    return `${alpha[c+ref_col]}`;
  });
  return cell;
}

let formatCellsList = [];

function formatCellsAdd(format, rstart, cstart, rend, cend) {
  // --- add formatting locations (row start, col start, row end, col end) to list --- //
  var alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split(''); // only 26 columns supported
  formatCellsList.push({
    'format': format,
    'range' : `${alpha[cstart]}${rstart}:${alpha[cend]}${rend}`
  }); // convert rc to A1 Notation
}

function formatCellsApply() {
  // --- apply formatting to added format cell locations --- //
  if ( formatCellsList.length ) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    for ( var i = 0; i < formatCellsList.length; i++ ) {
      var format = formatCellsList[i].format;
      var range = sheet.getRange(formatCellsList[i].range);
      if ( format == 'bold' ) {
        range.setFontWeight('bold');
      }
      if ( format == 'blackHeader' ) {
        range.setFontWeight('bold');
        range.setBackground('black');
        range.setFontColor('white');
      }
      if ( format == 'lightHeader' ) {
        range.setFontWeight('bold');
        range.setBackground('#efefef');
      }
    }
  }
}

function renderCSVList(rows, colLimit = 8) {
  // --- receives a 2 dimensional list (csv style) (e.g. [['', ''], ['', '']]) & renders to spreadsheet --- //
  var rowCount = rows.length; // number of rows in the CSV
  var colCount = Math.min(colLimit, rows[0].length); // number of columns with limit
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(1, 1, rowCount, colCount); // set values in the range A1:colLimit(dynamic)
  sheet.getRange(1, 1, rowCount + 100, colCount).clearContent(); // clear values (+100 of last row)
  sheet.getRange(1, 1, rowCount + 100, colCount).clearFormat();  // clear formatting (+100 of last row)
  range.setValues(rows);
  formatCellsApply();
}

// -- End: Sheet Cell Helpers -- //

// -- Workout Data Raw Format Parser & CSV Generator Main: workoutDataparser() -- //

function workoutDataParser(input) {
  // --- receives a single workout data entry (v1 or v2) and converts it into object --- //
  let workout_data = { entry_type: '', workout_details: {}, workout_list: [] };

  // parser start
  let lines = input.split('\n');
  if ( lines.length ) {
    let header = lines[0];

    let workout_attrs = {};
    let global_set_count = '';
    let global_rest_time = '';
    
    // section for v1 workouts format (e.g. workouts 4x rst:30s) (must start with "workouts")
    if ( header.slice(0, 8) == 'workouts' ) { // "workouts" plural not singular v1
      // header row parser start
      let head_cols = header.split(' ');
      if ( head_cols.length > 1 ) {
        // order of columns is: "workouts col1 col2 etc."
        // - global_set_count must be in col1 or won't be parsed
        // - global_rest_time can be in col col2 etc.
        for ( var i = 1; i < head_cols.length; i++ ) {
          let cur_col = head_cols[i];
          // global_set_count parser
          if ( i == 1 && cur_col[cur_col.length-1] == 'x' ) {
            global_set_count = cur_col.replace(/[a-zA-Z]/g, '');
          }
          // global_rest_time parser
          else if ( cur_col.slice(0, 4) == 'rst:' ) {
            global_rest_time = cur_col.replace(/^rst\:/g, '');
          }
        }
        workout_data.entry_type = 'workout';
        workout_data.workout_details = {
          name     : '',
          date     : '',
          hour     : '',
          dur      : '',  // total workout duration
          worktime : '',  // total non-rest/work time
          garmin   : '',  // garmin url to workout (optional)
          global_set_count: global_set_count,
          global_rest_time: global_rest_time,
        }
      }
      // header row parser end

      // body rows parser start
      for ( var i = 1; i < lines.length; i ++ ) {
        let cur_line = lines[i];
        // each row or line must start with either "-" or "."
        if ( cur_line[0] == '-' || cur_line[0] == '.' ) {
          cur_line = cur_line.replace(/^[-\.]\s+/g, '');
          let cur_split  = cur_line.split(',');
          let cur_name   = cur_split[0].trim();
          let cur_reps   = cur_split[1].trim();
          let cur_sets   = '';
          let cur_weight = '';

          // check if names end with ("lb" or "kg")
          let cur_name_ending = cur_name[cur_name.length-2] + cur_name[cur_name.length-1];
          if ( cur_name_ending == 'lb' || cur_name_ending == 'kg' ) {
            let cur_name_split = splitAtLast(cur_name, ' ');
            cur_name = cur_name_split[0].trim();
            cur_weight = cur_name_split[1].trim();
          }

          // check if name starts with set count (e.g. "4x")
          if ( new RegExp(/^[\d]{1,2}x/g).test(cur_name) ) {
            let cur_name_split = cur_name.split(/\s(.*)/s);
            cur_sets = cur_name_split[0].trim().replace(/[a-zA-Z]/g, '');
            cur_name = cur_name_split[1].trim();
          }

          let cur_data = {
            name   : cur_name,
            sets   : cur_sets || global_set_count,
            reps   : cur_reps,
            weight : cur_weight,
            rest   : global_rest_time,
          }
          workout_data.workout_list.push(cur_data);
        }
      }
      // body rows parser end
    }
    // end section for v1 workouts format
    
    // section for v2 workouts format (e.g. workout, d/m hr, ({name}), ({time=00:00,key=val,}))
    else if ( header.slice(0, 7) == 'workout' ) { // "workout" singular not plural v2
      let head_cols = header.split(/,(?![^(]*\))/); // split at commas not in parenthesis
      if ( head_cols.length > 1 ) {
        let date_hour = head_cols[1].trim().split(' ');
        let workout_date = date_hour[0];
        let workout_hour = date_hour[1];
        let workout_name = head_cols[2].trim().replace(/^\(|\)$/g, '');
        let key_val_meta = function(str) {
          str = str.replace(/^\(|\)$/g, '');
          let obj = str.split(',');
          let fin = {};
          for ( var i = 0; i < obj.length; i++ ) {
            let objspl = obj[i].trim().split('=');
            fin[ objspl[0].trim() ]  = objspl[1].trim();
          }
          return fin;
        };
        let workout_meta = head_cols[3] ? key_val_meta(head_cols[3].trim()) : {};
        global_set_count = workout_meta['sets'] ? workout_meta['sets'] : '';
        global_rest_time = workout_meta['rst'] || '';

        workout_data.entry_type = 'workout';
        workout_data.workout_details = {
          name     : workout_name,
          date     : workout_date,
          hour     : workout_hour,
          dur      : workout_meta['dur'] || '',      // total workout duration
          worktime : workout_meta['worktime'] || '', // total non-rest/work time
          garmin   : workout_meta['garmin'] || '',   // garmin url to workout (optional)
          global_set_count: global_set_count, // global set count
          global_rest_time: global_rest_time, // global rest time
        }
      }

      // body rows parser start
      for ( var i = 1; i < lines.length; i ++ ) {
        let cur_line = lines[i];
        // each row or line must start with either "-" or "."
        if ( cur_line[0] == '-' || cur_line[0] == '.' ) {
          cur_line = cur_line.replace(/^[-\.]\s+/g, '');
          let cur_split  = cur_line.split(',');
          let cur_sets   = cur_split[0].trim().replace(/[a-zA-Z]/g, '');
          let cur_name   = cur_split[1].trim();
          let cur_weight = cur_split[2].trim();
          let cur_reps   = cur_split[3] ? cur_split[3].trim() : '';
          let cur_rest   = cur_split[4] ? cur_split[4].trim() : '';
          let cur_data = {
            name   : cur_name,
            sets   : cur_sets || global_set_count,
            reps   : cur_reps,
            weight : cur_weight,
            rest   : cur_rest || global_rest_time,
          }
          workout_data.workout_list.push(cur_data);
        }
      }
      // body rows parser end
    }
    // end section for v2 workouts format

    // section for optional headers
    else if ( header.toLowerCase().includes('week') ) {
      let week_val = header.toLowerCase().match(/week\s([1-6])/);
      week_val = week_val ? 'Week ' + week_val[1] : 'Week';
      workout_data['entry_type'] = 'week';
      workout_data['workout_details']['name'] = week_val;
    }
    // end section for optional headers
  }
  return workout_data;
}

function parseWorkouts(filedata) {
  // --- receives a string of multiple workouts (typically for a month or more) and runs parser on each --- //
  let final = [];
  filedata = filedata.split(/\n\s*\n/).filter(item=>item);
  if ( filedata.length ) {
    final.push(['Date','Sets','Exercise','Weights','Reps','Rep Avg','Rep Tot','Hr/Rest Set:Ex']);
    formatCellsAdd('blackHeader', final.length, 0, final.length, 7); // set blackHeader for main header row
    for ( var i = 0; i < filedata.length; i++ ) {
      if ( filedata[i] ) {
        let workout = workoutDataParser(filedata[i].trim());
        if ( workout['entry_type'] == 'workout' ) {
          let wkdetails = workout['workout_details'];
          let wklist = workout['workout_list'];
          let calc_tot_down = 
            `=sum(` + 
              `$cellref(1,0):INDEX(` + 
                `$cellref(1,0):$cellref(100,0),` + 
                `MATCH(TRUE,($cellref(1,0):$cellref(100,0)=""),0)` + 
              `)` +
            `)`;
          final.push([convertDate(wkdetails['date']), calc_tot_down, wkdetails['name'], '', '', '', calc_tot_down, wkdetails['hour']]);
          formatCellsAdd('bold', final.length, 0, final.length, 7); // set bold for 'date' rows
          for ( var j = 0; j < wklist.length; j++ ) {
            let wkeach = wklist[j];
            let calc_rep_avg = '=iferror(round(average(split($cellref(0,-1)," ")), 2))';
            let calc_rep_tot = '=iferror(sum(split($cellref(0,-2)," ")))';
            final.push(['', wkeach['sets'], wkeach['name'], wkeach['weight'], wkeach['reps'], calc_rep_avg, calc_rep_tot, wkeach['rest']]);
          }
          final.push(['', '', '', '', '', '', '', '']);
        }
        
        else if ( workout['entry_type'] == 'week' ) {
          let calc_sum_week_sets = 
            `=sum(` + 
              `$cellref(1,0):INDEX(`+
                `$cellref(1,0):$colref(0),`+
                `IFERROR(MATCH("Week*",$cellref(1,-1):$colref(-1),0)-1, 1000)`+
              `)`+
            `)/2`;
          let calc_sum_week_reps = 
            `=sum(` + 
              `$cellref(1,0):INDEX(`+
                `$cellref(1,0):$colref(0),`+
                `IFERROR(MATCH("Week*",$cellref(1,-6):$colref(-6),0)-1, 1000)`+
              `)`+
            `)/2`;
          final.push([workout['workout_details']['name'], calc_sum_week_sets, '', '', '', '', calc_sum_week_reps, '']);
          formatCellsAdd('lightHeader', final.length, 0, final.length, 7); // set lightHeader for 'week' rows
        }
      }
    }
  }

  for ( var r = 0; r < final.length; r++ ) {
    for ( var c = 0; c < final[r].length; c++ ) {
      final[r][c] = convertCellLocationTags(final[r][c], r, c);
    }
  }
  
  return final;

}

// --- App Start: populateWorkouts(), onEdit(e) -- //

function populateWorkouts() {
  // --- main function that calls everything --- //
  const cell_data    = 'J1';
  const cell_remote  = 'J2';
  const cell_udate   = 'L1';
  const data_columns = 8; // (A:H)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let remote_options = { 'use' : 'false', 'url': '', 'hash' : '' };
  if ( sheet.getRange(cell_remote).getValue().toLowerCase() == 'remote' ) {
    let remote_note = sheet.getRange(cell_remote).getNote();
    if ( remote_note.length ) {
      let matches = [...remote_note.matchAll(/(use|url|hash):\s*([\S]+)/gi)];
      if ( matches.length ) {
        for ( var i = 0; i < matches.length; i++ ) {
          remote_options[ matches[i][1].toLowerCase() ] = matches[i][2];
        }
      }
    }
  }

  let parsed = [];
  if ( remote_options.use.toLowerCase() == 'true' ) {
    parsed = parseWorkouts(embedRemoteData(remote_options.url, remote_options.hash));
  } else if ( sheet.getRange(cell_data).getValue().toLowerCase() == 'data' ) {
    parsed = parseWorkouts(sheet.getRange(cell_data).getNote());
  }

  if ( parsed.length ) {
    renderCSVList(parsed, data_columns);
    sheet.getRange(cell_udate).setValue('Upd: ' + currentTime());
    sheet.getRange(cell_udate).setNote('Updated: ' + currentTime('long'));
  } else {
    alrt('No usable data source found. Please supply a URL or cell note data.');
  }
}

// --  Start: Check Box Trigger -- //

function onEdit(e) {
  // --- checkbox button actions for non-web/mobile devices --- //
  let cell_chkbx = 'K1';
  if ( e.range.getA1Notation() === cell_chkbx && ( e.value === "TRUE" || e.value === "FALSE" ) ) {
    populateWorkouts();
  }
}

// -- End: Check Box Trigger -- //

// -- End: App Start -- //

// === End: Workout Manager Main === //




/*

Data Examples

--- v1 data example ---
workouts 4x rst:30s
- 5x pull up, 5 3 2 2
- db ovh press 2x20lb, 20 12 12 10

--- v2 data example ---
workout, 1/11 7p, (sw1: pull ups, 1 biceps), (garmin=id_or_url, key=val)
. 4, pull up, body, 7 4 3 4, 45s
. 4, hammer, 2x25lb, 16 12 8 10, 30s

*/

let v1 = 'workouts 4x rst:30s\n' +
          '- pull up, 5 3 2 2\n'+ 
          '- 3x db ovh press 2x20lb, 20 12 12 10';

let v2 = 'workout, 1/11 7p, (sw1: pull ups, 1 biceps), (garmin=hey, other=hi)\n' + 
          '. 4, pull up, body, 7 4 3 4, 45s\n'+
          '. 4, hammer, 2x25lb, 16 12 8 10, 30s';

let dataTester = function() {
  console.log(JSON.stringify(workoutDataParser(v1), '', 2));
  console.log(JSON.stringify(workoutDataParser(v2), '', 2));
}



