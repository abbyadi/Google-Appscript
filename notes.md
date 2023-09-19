
**geoRegionSort Function Description**
* Sets up a fund direction based on the funding Round (R1/R2). Smaller to larger county in R1 and vice-versa in R2.
* Then loop through each region to set headers and sheet data. Previously funded projects in set-asides are labelled "Fund S/A"
* Number of Cycles to loop through each region is determined by extracting number of unfunded deals. If the number of un-funded deals in a region is more than the previous regions then that number is set as the number of cycles (assigend to numCycles).
