## Workout Manager (Apps Script)

<!-- git repo create: 1/18/2025 -->

**Google Apps Script Project Structure**
> Code.gs

**Testing**

```javascript
let v1 = 'workouts 4x rst:30s\n' +
          '- pull up, 9 8 7 7\n'+ 
          '- 3x db ovh press 2x30lb, 20 12 12 10';

let v2 = 'workout, 1/11 7p, (sw1: pull ups, 1 biceps), (garmin=id_or_url, key=val)\n' + 
          '. 4, pull up, body, 12 10 10 9, 45s\n'+
          '. 4, curl, 2x25lb, 16 12 8 10, 30s';

let main = function() {
  console.log(JSON.stringify(workoutDataParser(v1), '', 2));
  console.log(JSON.stringify(workoutDataParser(v2), '', 2));
}

```

**Running:** via Google Spreadsheet

```javascript
populateWorkouts()
```