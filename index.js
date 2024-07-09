#! /usr/bin/env node
const { program } = require('commander')
const homedir = require('os').homedir();
const azdev = require("azure-devops-node-api");
const fs = require('fs')
const joinPath = require('path.join');
const chalk = require('chalk');
const config = require('config');
var downloadPath = joinPath(homedir,'Documents');

program.command("download")
        .description("Download All Attachments from Work Item provided it's ID")    
        .argument("<WorkItemId>", "The work item id to get attachments from")
        .option('-d,--downloadPath [DownloadPath]','The path to download the attachments',downloadPath)
        .action(async function download(WorkItemId,options) {                
            console.log(chalk.green("Establishing Azure Devops Server Connection..."))
            //Init connection
            let orgUrl = config.get('azureServer.BaseUrl');
            //let token = "o6e6i5u7gccxrujfjqlsm2fxws2yoap2xxo2viy2yn4bqfzrnrpq";       
            let token = config.get('azureServer.Token');   
            let authHandler = azdev.getPersonalAccessTokenHandler(token); 
            let option = {
                 cert: {
                     caFile: config.get('azureServer.CertLocation')
                 },
             };
            
            try{
                let connection = new azdev.WebApi(orgUrl, authHandler,option);   
                console.log(chalk.green("Getting Item Details for "+WorkItemId))
                //Getting Work Item Details
                let witApi  = await connection.getWorkItemTrackingApi();
                let workItem = await witApi.getWorkItem(WorkItemId,undefined,undefined,4)
                var folderPath = joinPath(options.downloadPath,(WorkItemId+"_"+workItem.fields["System.Title"]).replace(/[^\w\s]/gi, ''));
                if (fs.existsSync(folderPath)) {
                    fs.rmdirSync(folderPath,{ recursive: true, force: true })
                  }
                fs.mkdirSync(folderPath);
                console.log(chalk.green("Downloading Attachments for "+WorkItemId))
                //Downloading Attachments
                workItem.relations.filter(rel => rel.rel === 'AttachedFile' ).forEach(async element => {     
                    let readableStream = await witApi.getAttachmentContent(element.url.split("/").at(-1),element.attributes['name'],undefined,true)
                    let writableStream = fs.createWriteStream(joinPath(folderPath,element.attributes['name']));
                    readableStream.pipe(writableStream);
                });
                console.log(chalk.green("Successfully download "+workItem.relations.filter(rel => rel.rel === 'AttachedFile' ).length+" attachments"))
            } catch (error) {
                console.error(error)
            }

            })

program.parse()

