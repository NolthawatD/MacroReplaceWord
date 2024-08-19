# README #

https://bitbucket.org/planet-x-technology/workspace/projects/DEAL
- backoffice-deal
- deal-backend

แตก branch จาก
backoffice-deal => to-develop => feature/macro-contract
deal-backend    => develop    => feature/macro-contract

# Installation #
* npm install or yarn
* copy .env.development and remove .development out

# Installation with use on docker #
* install docker desktop in your computer        //https://docs.docker.com/desktop/install/windows-install/
* open project on vscode

- start
* docker --version                               // check existing docker 
* docker login                                   // for check auth, can sign in with docker desktop, it's easier
* docker build -t macro-service-api .            // build image
* docker run -p 8080:8080 -d macro-service-api   // proceeding on container.  docker run -d -p 8080:8080 -v /tmp:/app/config macro-service-api

- stop
* docker container ls                            // check container proceeding
* docker container stop 15dd41354e69             // stop container by tag id

# Run Command #
* node src/. or node src/index.js 

# git repository #
* https://bitbucket.org/planet-x-technology/macro-service-api/src/develop/


# Docker #
docker build -t macro-service-api  .        
docker run -p 8080:8080 -d macro-service-api
docker container 
docker container stop 15dd41354e69

# Docker Command Push #
https://medium.com/allenhwkim/getting-started-docker-on-mac-os-x-72c64670464a#id_token=eyJhbGciOiJSUzI1NiIsImtpZCI6ImVkODA2ZjE4NDJiNTg4MDU0YjE4YjY2OWRkMWEwOWE0ZjM2N2FmYzQiLCJ0eXAiOiJKV1QifQ.eyJpc3MiOiJodHRwczovL2FjY291bnRzLmdvb2dsZS5jb20iLCJhenAiOiIyMTYyOTYwMzU4MzQtazFrNnFlMDYwczJ0cDJhMmphbTRsamRjbXMwMHN0dGcuYXBwcy5nb29nbGV1c2VyY29udGVudC5jb20iLCJhdWQiOiIyMTYyOTYwMzU4MzQtazFrNnFlMDYwczJ0cDJhMmphbTRsamRjbXMwMHN0dGcuYXBwcy5nb29nbGV1c2VyY29udGVudC5jb20iLCJzdWIiOiIxMTA1Njg2MzM3ODUzMzE3NTQ3NTEiLCJlbWFpbCI6Im5vbHRoYXdhdC5kQGdtYWlsLmNvbSIsImVtYWlsX3ZlcmlmaWVkIjp0cnVlLCJuYmYiOjE3MDgxNzk2NjUsIm5hbWUiOiJOb2x0aGF3YXQgRGVhbmdoYXJuIiwicGljdHVyZSI6Imh0dHBzOi8vbGgzLmdvb2dsZXVzZXJjb250ZW50LmNvbS9hL0FDZzhvY0pNcC1hMTd1TkIwUURwaTNEM2UxVEFuU1oxNHR1MVNFMnVQSUZjbmRyZD1zOTYtYyIsImdpdmVuX25hbWUiOiJOb2x0aGF3YXQiLCJmYW1pbHlfbmFtZSI6IkRlYW5naGFybiIsImxvY2FsZSI6ImVuIiwiaWF0IjoxNzA4MTc5OTY1LCJleHAiOjE3MDgxODM1NjUsImp0aSI6IjM5Nzk2MzVkYjRjNzRiM2FhNzU3MjM0YTEwNGJkYjU4OTNlZGZlYWUifQ.nO8DslrxETi3SndMtpwhL-h9JhFxf10eBVOGtMM3fETHGZkMT_I8iDdmytzm0ulKxRAWMKb3Dw6cj5aimkV8E9zfDZqq0untfIAGFHrN7fIJ4bqAFiDDRkk2M74RobkZWFK5_3Qy4Stdw1nq2taT8k8pjAVY4L6vowSwi6l81xkJ8fldGFhpu5_uZRoYiJj-rHipgVekSFT3ylRgM4oIBNV_WNJj0vQ-xblayNE0GbVywOh2Qmb15LJzwj1K-hF4-nAmu0A6PK86MaINLs9UKsd9uLhGmMekvzskW5QAaZX4I7XLSsesN1jEQ3usmcL4_Huq4pqrAAvdGEpyuWxDaw

# Docker Command Pull #
https://www.stacksimplify.com/aws-eks/docker-basics/get-docker-image-from-docker-hub-and-run-/

This README would normally document whatever steps are necessary to get your application up and running.

### What is this repository for? ###

* Quick summary
* Version
* [Learn Markdown](https://bitbucket.org/tutorials/markdowndemo)

### How do I get set up? ###

* Summary of set up
* Configuration
* Dependencies
* Database configuration
* How to run tests
* Deployment instructions

### Contribution guidelines ###

* Writing tests
* Code review
* Other guidelines

### Who do I talk to? ###

* Repo owner or admin
* Other community or team contact

# google-doc-sheet-api
for poc job thai id

# video-youtube
https://www.youtube.com/watch?v=PFJNJQCU_lo //credential sheet
https://www.youtube.com/watch?v=0KoZSVnTnkA //credential doc


# google api library
https://console.cloud.google.com/apis/library/browse?project=mindful-genius-413205&q=google%20docs%20api

# email google for approve from owner google sheet
google-sheet@mindful-genius-413205.iam.gserviceaccount.com

# get credentials key
https://console.cloud.google.com/iam-admin/serviceaccounts/details/107244488346950868602;edit=true/keys?project=mindful-genius-413205

# step 
1. npm init
2. npm install express googleapis

# git step
git checkout -b ＜new-branch＞
git add .
git commit -m 'google-doc'

https://www.atlassian.com/git/tutorials/using-branches/git-merge

# Start a new feature
git checkout -b new-feature main

# Edit some files
git add <file>
git commit -m "Start a feature"

# Edit some files
git add <file>
git commit -m "Finish a feature"
git push origin develop > username , password

# Merge in the new-feature branch
git checkout main
git merge new-feature
git branch -d new-feature

# Prompt Generate SwaggerDoc
Please write JSDoc comment with swagger annotations 

# Add Folder Config
1. add credential doc, sheet, token (generate form veify first)


# data to sending about demo1
{
  "documentId": "1IbS-iVIpNbKRke6_fxb5lMwodMj8muj1123D8e3q6Zk",
  "sheets": [
    {
      "sheetId": "16cZb_KaasnIXJlZNTaijGDrREpPMRGQGlMiC84aJZc8",
      "sheetName": "Investor",
      "columnFormat": "G",
      "columnsName": []
    },
    {
      "sheetId": "16cZb_KaasnIXJlZNTaijGDrREpPMRGQGlMiC84aJZc8",
      "sheetName": "SME",
      "columnFormat": "G",
      "columnsName": []
    }
  ],
  "fromRow": 10,
  "toRow": 10
}
