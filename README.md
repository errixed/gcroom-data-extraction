# [![My Skills](https://skills.thijs.gg/icons?i=js)](https://skills.thijs.gg) groom-data-extraction
This code gives you with the submission status of all students in all assignments in the selected course along with their name and email in the google sheets.
<br/>
<br/>
![GitHub followers](https://img.shields.io/github/followers/errixed)
![GitHub forks](https://img.shields.io/github/forks/errixed/groom-data-extraction)
![GitHub watchers](https://img.shields.io/github/watchers/errixed/groom-data-extraction)
![GitHub Repo stars](https://img.shields.io/github/stars/errixed/groom-data-extraction)
## Services

<img
src="https://www.gstatic.com/images/branding/product/2x/sheets_96dp.png"
align="left"
width="70px"/>
### sheets

<img
src="https://www.gstatic.com/images/branding/product/2x/classroom_96dp.png"
align="left"
width="70px"/>
### classroom

<br/>

## Setup
1. go to `Google Drive`
2. create a `new Google Sheets`
3. in the `Extensions` tab, select `Apps Script` (Apps Script will be opened)
4. in Apps Script, go to `Settings` and check the `Show "appsscript.json" manifest file in editor` option
5. copy and paste codes in `appsscript.json` from this repo to apps script
6. copy and paste codes in `Code.gs` from this repo to apps script
7. replace `SELECTED_COURSE_ID` on line 26 with `your course ID`

## On Run
Run courseData function to get course IDs that you have access with
Run assignmentSubmissionState function to get students data (names, emails, submission statuses)
