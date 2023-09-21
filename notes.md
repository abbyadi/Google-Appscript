# Sort Notes

## geoRegionSort Function Description

- Sets up a fund direction based on the funding Round (R1/R2). Smaller to larger county in R1 and vice-versa in R2.
- The function then loops through each region to set headers and sheet data. Previously funded projects in set-asides are labelled "Fund S/A"
- Number of Cycles to loop through each region is determined by extracting number of unfunded deals and set to variable named numCycle. If the number of un-funded deals in a region is more than the previous regions then that number is set as the number of cycles (assigend to numCycles).

    ``` js
    let unfunded = dataRangeValues.slice(5).filter(row => !row[lastCol-1]); // filtering number of projects that are unfunded. slice(5) removes first 5 rows
    if(numCycles < unfunded.length) {
      numCycles = unfunded.length;
      console.log(numCycles)
    }
    ```

- The first step inside the while loop checks if it is the first cycle and if it is then the first project which is the highest scoring and ranking project will get funded as long it is not funded in a set-aside.

    ``` js
    if (runCycle === 0) { /**First Round of funding deals. Fund housing type even if negative if highest tiebreaker and first to get funded*/
        ...
    }
    ```

- Starting with second cycle: first step checks if any project has been skipped 'skip 125' and if so then the next project will only be funded if its score is equal to the first project that has been assigned 'skip 125' and if its tie-breaker sccore is at least 75% of the first 'skip 125' project. If not project is assigned 'skip 75%TB'.

    ``` js
    if (!dataRangeValues[i][7] >= 0.75*frstSkip125[7] || !dataRangeValues[i][6] >= frstSkip125[6]) {//project Tiebreaker is not 75% of 1st Skip 125 project TB or point score is not equal or greater than 1st Skip 125 project
        ...
    }
    ```

    If conditions are met the project is funded or skipped using the **fundGeo** function.
  - If there are no 'skip 125' projects the project is funded or skipped using the **fundGeo** function.

    ``` js
    else {
            fundGeo(i); //Fund current project at row i;
            if (breakSwitch === "on") {
            break;
            }
        }
    ```

- ### Description for fundGeo closure function

    The fundGeo function is subset of the geoRegionSort function and forms a closure.

- The function checks if the project that has been passed to it as an argument passes or fails four test:

```js
        if (!dataRangeValues[currP][lastCol-1]) { //Checks that Project is not Funded or Skipped in the Fund/Skip Column
          if (dataRangeValues[currP][6] >= minScr) { //Project meets or exceeds min score.
            if(dataRangeValues[currP][5]<=balanceAmt){
              if(regionCounter.get(element) > 0) { /**checks if region has run out of Credits.*/
                if (dataRangeValues[currP][6]>=dataRangeValues[currP-1][6]) {
                    ...
                }
              }
            }
          }
        }
```
