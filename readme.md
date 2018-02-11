# How To do excel operations in NodeJs (in-memory)

In this example you can easily create a excel template using a custom defined JSON input with support of various validators on columns. And the without need to save this excel into external space you can pipe it to response directly. I'm using **excel4node** for generating excel.
Second functionality allows you to upload excel and get a parsed JSON of your file and again you don't need any extra external space to save incoming file. I'm using **fileupload** and **xlsx modules**.


## Files

Steps to run
1) clone the repo
2) run npm install
3) npm start