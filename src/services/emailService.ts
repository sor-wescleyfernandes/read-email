// import { logger } from "../utils/logger";
import { simpleParser } from "mailparser";
import Imap from "node-imap";

// Configurações para acessar o servidor IMAP do Outlook
const config: Imap.Config = {
  user: "wescley.fernandes@hotmail.com",
  password: "Amanh@01",
  host: "outlook.office365.com",
  port: 993,
  tls: true
};

const fs = require('fs');
const xlsx = require('xlsx');
const moment = require('moment');
const currentDate = moment().subtract(1, 'days').format('MMMM D, YYYY');

export const readEmails = () => {
  const imap = new Imap(config);

  imap.connect();

  imap.once("ready", () => {
    imap.openBox("INBOX", true, (err: any, box: any) => {
      if (err) throw err;
      imap.search(["UNSEEN", ["SINCE", currentDate]], (err: any, results: any) => {
        if (err) throw err;
        const f = imap.fetch(results, { bodies: "" });
        f.on("message", (msg: Imap.ImapMessage, seqno: number) => {
            msg.on("body", (stream: any, info: any) => {
                simpleParser(stream, {}, (err: any, parsed: any) => {
                    if (err) throw err;

                    if (parsed.attachments.length > 0) {
                      parsed.attachments.forEach((attachment: { contentType: string; filename: any; content: any; }) => {
                          if (attachment.contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
                              fs.writeFileSync(`./${attachment.filename}`, attachment.content);
                              const workbook = xlsx.readFile(`./${attachment.filename}`);
                              const sheet_name_list = workbook.SheetNames;
                              sheet_name_list.forEach(function(y: string | number) {
                                  var worksheet = workbook.Sheets[y];
                                  var headers: { [key: string]: any } = {};
                                  var data: { [key: number]: { [key: string]: any } } = [];
                                  for(const z in worksheet) {
                                      if(z[0] === '!') continue;
                                      //parse out the column, row, and value
                                      var tt = 0;
                                      for (var i = 0; i < z.length; i++) {
                                          if (!isNaN(parseInt(z[i]))) {
                                              tt = i;
                                              break;
                                          }
                                      };
                                      var col = z.substring(0,tt);
                                      var row = parseInt(z.substring(tt));
                                      var value = worksheet[z].v;

                                      //store header names
                                      if(row == 1 && value) {
                                          headers[col] = value;
                                          continue;
                                      }

                                      if(!data[row]) data[row]={};
                                      data[row][headers[col]] = value;
                                  }
                                  //drop those first two rows which are empty
                                  // data.shift();
                                  // data.shift();
                                  console.log(data);
                              });
                          }
                      });
                  }

                    
                    // logger.info(`Subject: ${parsed.subject}`);
                    // logger.info(`Attachments: ${parsed.attachments}`);
                });
            });
        });
        // f.on("message", (msg: { on: (arg0: string, arg1: (stream: any, info: any) => void) => void; }, seqno: any) => {
        //   msg.on("body", (stream, info) => {
        //     simpleParser(stream, {}, (err: any, parsed: any) => {
        //       if (err) throw err;
        //     //   logger.info(`Subject: ${parsed.subject}`);
        //     //   logger.info(`Attachments: ${parsed.attachments}`);
        //     });
        //   });
        // });
        f.once("end", () => {
          imap.end();
        });
      });
    });
  });

  imap.once("error", (err: any) => {
    debugger;
    console.log(err);
  });
};
