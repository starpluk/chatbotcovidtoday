var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1zFf6JUOQTVxmqd-H9HhP9b0kt5II_zKXsKKLQ6Rb_AE/edit");
var sheet = ss.getSheetByName("World");
var sheet_2 = ss.getSheetByName("Thai");

function doPost(e) {

    var data = JSON.parse(e.postData.contents)
    var userMsg = data.originalDetectIntentRequest.payload.data.message.text;
    var values = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    userMsg = userMsg.toLowerCase();
    if (userMsg == "korea") { userMsg = "s. korea" }

    if (userMsg == "Thailand" || userMsg == "ไทย" || userMsg == "ประเทศไทย") {
        var total_th = sheet_2.getRange(2, 1).getValue();
        var active_th = sheet_2.getRange(2, 3).getValue();
        var recover_th = sheet_2.getRange(2, 2).getValue();
        var deaths_th = sheet_2.getRange(2, 4).getValue();
        var datetoday_th = sheet_2.getRange(1, 10).getValue();

        var result = {
            "fulfillmentMessages": [
                {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line": {
                            "altText": "Flex Message",
                            "contents": {
                                "type": "bubble",
                                "body": {
                                    "type": "box",
                                    "contents": [
                                        {
                                            "type": "box",
                                            "contents": [
                                                {
                                                    "size": "xl",
                                                    "weight": "bold",
                                                    "color": "#E37E7E",
                                                    "text": "Covid-19 Thailand",
                                                    "type": "text"
                                                },
                                                {
                                                    "url": "https://s3-ap-southeast-1.amazonaws.com/img-in-th/700834c41c8c13879034341be28826aa.png",
                                                    "size": "xl",
                                                    "type": "icon"
                                                }
                                            ],
                                            "layout": "baseline"
                                        },
                                        {
                                            "spacing": "sm",
                                            "type": "box",
                                            "layout": "vertical",
                                            "contents": [
                                                {
                                                    "contents": [
                                                        {
                                                            "weight": "bold",
                                                            "type": "text",
                                                            "margin": "sm",
                                                            "align": "start",
                                                            "gravity": "center",
                                                            "text": "ผู้ติดเชื้อ",
                                                            "size": "xxl"
                                                        },
                                                        {
                                                            "align": "end",
                                                            "weight": "bold",
                                                            "size": "4xl",
                                                            "gravity": "center",
                                                            "color": "#E83D3D",
                                                            "type": "text",
                                                            "text": total_th
                                                        }
                                                    ],
                                                    "type": "box",
                                                    "layout": "baseline"
                                                },
                                                {
                                                    "type": "text",
                                                    "text": "-----------------------------------"
                                                },
                                                {
                                                    "contents": [
                                                        {
                                                            "size": "md",
                                                            "text": "กำลังรักษา",
                                                            "weight": "bold",
                                                            "margin": "sm",
                                                            "type": "text"
                                                        },
                                                        {
                                                            "text": active_th,
                                                            "weight": "bold",
                                                            "align": "end",
                                                            "size": "md",
                                                            "color": "#EAC919",
                                                            "type": "text"
                                                        }
                                                    ],
                                                    "layout": "baseline",
                                                    "type": "box"
                                                },
                                                {
                                                    "type": "box",
                                                    "layout": "baseline",
                                                    "contents": [
                                                        {
                                                            "size": "md",
                                                            "text": "หายแล้ว",
                                                            "type": "text",
                                                            "weight": "bold"
                                                        },
                                                        {
                                                            "type": "text",
                                                            "size": "md",
                                                            "color": "#47CC08",
                                                            "align": "end",
                                                            "weight": "bold",
                                                            "text": recover_th
                                                        }
                                                    ]
                                                },
                                                {
                                                    "contents": [
                                                        {
                                                            "type": "text",
                                                            "weight": "bold",
                                                            "text": "เสียชีวิต",
                                                            "size": "md"
                                                        },
                                                        {
                                                            "size": "md",
                                                            "type": "text",
                                                            "color": "#726767",
                                                            "text": deaths_th,
                                                            "align": "end",
                                                            "weight": "bold"
                                                        }
                                                    ],
                                                    "type": "box",
                                                    "layout": "baseline"
                                                }
                                            ]
                                        },
                                        {
                                            "type": "text",
                                            "color": "#AAAAAA",
                                            "size": "sm",
                                            "text": datetoday_th,
                                            "wrap": true
                                        }
                                    ],
                                    "layout": "vertical",
                                    "spacing": "md"
                                },
                                "footer": {
                                    "type": "box",
                                    "contents": [
                                        {
                                            "size": "md",
                                            "type": "spacer"
                                        },
                                        {
                                            "type": "button",
                                            "action": {
                                                "uri": "https://covid19.workpointnews.com/",
                                                "label": "Website",
                                                "type": "uri"
                                            },
                                            "color": "#0DA11A",
                                            "style": "primary"
                                        }
                                    ],
                                    "layout": "vertical"
                                }
                            },
                            "type": "flex"
                        }
                    }
                }
            ]
        }

    }


    else {
        for (var i = 0; i < values.length; i++) {
            var text = values[i][0];
            text = text.toLowerCase();
            var nameCountry = values[i][0];


            if (text == userMsg) {
                i = i + 2;
                var total = sheet.getRange(i, 2).getValue();
                var active = sheet.getRange(i, 7).getValue();
                var recover = sheet.getRange(i, 6).getValue();
                var deaths = sheet.getRange(i, 10).getValue();
                var day = sheet.getRange(1, 11).getValue();
                var mount = sheet.getRange(1, 12).getValue();
                var year = sheet.getRange(1, 13).getValue();
                var datetoday = sheet.getRange(1, 13).getValue();

                var result = {
                    "fulfillmentMessages": [
                        {
                            "platform": "line",
                            "type": 4,
                            "payload": {
                                "line": {



                                    "type": "flex",
                                    "altText": "Flex Message",
                                    "contents": {
                                        "type": "bubble",
                                        "body": {
                                            "type": "box",
                                            "layout": "vertical",
                                            "spacing": "md",
                                            "action": {
                                                "type": "uri",
                                                "label": "Action",
                                                "uri": "https://linecorp.com"
                                            },
                                            "contents": [
                                                {
                                                    "type": "text",
                                                    "text": "Covid-19 " + nameCountry,
                                                    "size": "xl",
                                                    "weight": "bold",
                                                    "color": "#E37E7E"
                                                },
                                                {
                                                    "type": "box",
                                                    "layout": "vertical",
                                                    "spacing": "sm",
                                                    "contents": [
                                                        {
                                                            "type": "box",
                                                            "layout": "baseline",
                                                            "contents": [
                                                                {
                                                                    "type": "text",
                                                                    "text": "ผู้ติดเชื้อ",
                                                                    "margin": "sm",
                                                                    "weight": "bold"
                                                                },
                                                                {
                                                                    "type": "text",
                                                                    "text": "" + total,
                                                                    "size": "lg",
                                                                    "align": "end",
                                                                    "weight": "bold",
                                                                    "color": "#FB4747"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "box",
                                                            "layout": "baseline",
                                                            "contents": [
                                                                {
                                                                    "type": "text",
                                                                    "text": "กำลังรักษา",
                                                                    "margin": "sm",
                                                                    "weight": "bold"
                                                                },
                                                                {
                                                                    "type": "text",
                                                                    "text": "" + active,
                                                                    "size": "lg",
                                                                    "align": "end",
                                                                    "weight": "bold",
                                                                    "color": "#EAC919"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "box",
                                                            "layout": "baseline",
                                                            "contents": [
                                                                {
                                                                    "type": "text",
                                                                    "text": "หายแล้ว",
                                                                    "margin": "sm",
                                                                    "gravity": "top",
                                                                    "weight": "bold"
                                                                },
                                                                {
                                                                    "type": "text",
                                                                    "text": "" + recover,
                                                                    "size": "lg",
                                                                    "align": "end",
                                                                    "weight": "bold",
                                                                    "color": "#47CC08"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "box",
                                                            "layout": "baseline",
                                                            "contents": [
                                                                {
                                                                    "type": "text",
                                                                    "text": "เสียชีวิต",
                                                                    "weight": "bold"
                                                                },
                                                                {
                                                                    "type": "text",
                                                                    "text": "" + deaths,
                                                                    "size": "lg",
                                                                    "align": "end",
                                                                    "weight": "bold",
                                                                    "color": "#726767"
                                                                }
                                                            ]
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "text",
                                                    "text": "ข้อมูลวันที่ " + datetoday,
                                                    "size": "xs",
                                                    "color": "#AAAAAA",
                                                    "wrap": true
                                                }
                                            ]
                                        },
                                        "footer": {
                                            "type": "box",
                                            "layout": "vertical",
                                            "contents": [
                                                {
                                                    "type": "spacer",
                                                    "size": "xxl"
                                                },
                                                {
                                                    "type": "button",
                                                    "action": {
                                                        "type": "uri",
                                                        "label": "Website",
                                                        "uri": "https://covid19.workpointnews.com/"
                                                    },
                                                    "color": "#0DA11A",
                                                    "style": "primary"
                                                }
                                            ]
                                        }
                                    }



                                }

                            }
                        }
                    ]
                }

            }
        }
    }
    var replyJSON = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    return replyJSON;
}