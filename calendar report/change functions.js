//SOQLでSalesforceから日報のデータを取得し、Excelにエクスポートする前に、変換処理をする。
var str = msg.payload; 
var astr = new Array();
for (i = 0; i < str.length; i++) {
    var obj = new Object();
    var p = eval('(' + JSON.stringify(str[i]) + ')');
    obj.件名 = (p.subject);
    obj.任命先 = (p.owner.Name);
    //時間格式変換
    p.activitydate = p.activitydate.replace(/-/g, "/");
    obj.日付 = (p.activitydate);
    if (p.isalldayevent == true) {//終日場合、開始時間nullなので、回避
        p.activitydatetime = p.activitydate
    }
    var date1 = new Date(p.activitydatetime);
    obj.開始時間 = (date1.toLocaleTimeString());
    obj.開始日時 = (date1.toLocaleString().replace(/-/g, "/"));
    var date2 = new Date(p.enddatetime);
    obj.終了時間 = (date2.toLocaleTimeString());
    obj.終了日時 = (date2.toLocaleString().replace(/-/g, "/"));
    obj.プロジェクト = (p.systemname__c);
    obj.勤務時間 = (p.worktime__c);//所要時間は休暇時間を含む可能性あるので、勤務時間も参考
    obj.理由 = (p.notices__c);
    obj.年 = (date1.getFullYear());
    obj.月 = (date1.getMonth() + 1);//月調整
    obj.曜日 = (date1.getDay());
    astr[i] = obj;
}

//月日数計算
//var date4 = new Date(astr[1].日付);
var yr = date1.getFullYear();  //年 of the last argument
var mon = date1.getMonth();    //月-1
var d = new Date(yr, mon + 1, 0);//今月の日数を取るため、月を来月にして、日を０にする。そうすれば、Dateオブジェクトは今月の最後の日になる。
var days = d.getDate();　　　　//将月份下移到下一个月份，同时将日期设置为0；由于Date里的日期是1~31，所以Date对象自动跳转到上一个月的最后一天；getDate（）获取天数即可。
var recds = new Array(days);
for (k = 0; k < days; k++) {
    var recd = new Object();
    recd.年 = yr;
    recd.月 = mon + 1;
    recd.日 = k + 1;
    recd.作業内容 = "";
    recd.理由 = "";
    var day4 = new Date(yr, mon, k + 1);
    var weekday = day4.getDay();
    switch (weekday % 7) {
        case 6:
            recd.曜日 = "土";
            break;
        case 1:
            recd.曜日 = "月";
            break;
        case 2:
            recd.曜日 = "火";
            break;
        case 3:
            recd.曜日 = "水";
            break;
        case 4:
            recd.曜日 = "木";
            break;
        case 5:
            recd.曜日 = "金";
            break;
        default:
            recd.曜日 = "日";
    }
    //開始、終了時間初期値
    var tempstartUTC = new Date(yr, mon, k + 1, 23, 59, 0);
    var tempstart = "";
    var tempendingUTC = new Date(yr, mon, k + 1, 0, 0, 1);
    var tempending = "";
    var n = 0;
    var events_repeat = new Array();
    var events_norepeat = new Array();
    var owner = "";
    var ownerstr = "";
    for (i = 0; i < astr.length; i++) {
        //日取得
        var date3 = new Date(astr[i].日付);
        var date = date3.getDate();
        //連日イベント場合、最初日の件名を取得必要なので当日はこのイベントの開始時間と終了時間の間に含まれてるかどうか判断する
        if (datecompare(date3, day4) && datecompare(day4, astr[i].終了日時)) {
            events_repeat[n] = astr[i].件名;
            n++;
            owner = astr[i].任命先;
            var temparr = unique(events_repeat);
            //前回のリストと一種場合、本ループの件名がリストに入ってるとを判明した上、のぞく。
            //また休日も出勤する場合、休日の件名を表示しない
            if (events_norepeat === undefined || temparr.length != events_norepeat.length) {
                events_norepeat = temparr;
                if(owner　==　"休日"){
                     if(ownerstr=="" || ownerstr.indexOf(owner) > -1){　//休日＋休日　　OK
                        recd.作業内容 += '●' + events_norepeat[events_norepeat.length - 1];
                     }else{　　　　　　　　　　　　　　　　　//勤務＋休日　　NG
                        ; 
                     }

                }else{
                    if(ownerstr=="" || ownerstr.indexOf(owner) > -1){　//勤務＋勤務　　OK
                        recd.作業内容 += '●' + events_norepeat[events_norepeat.length - 1];
                    }else{　　　　　　　　　　　　　　　　　//休日+勤務　　NG
                        recd.作業内容 =  '●' +astr[i].件名;
                    }
                }
                
            }
            ownerstr = ownerstr+owner;
            if (astr[i].理由 != null) {　　//理由の重複を無視された。
                recd.理由 += astr[i].理由;
            }
            //対象内のレコードのみ
            if (astr[i].プロジェクト != "休暇" && astr[i].プロジェクト != "AM休暇" && astr[i].プロジェクト != "PM休暇" && astr[i].プロジェクト != "休憩" && owner != "休日") {
                //開始時間
                if (Date.parse(new Date(astr[i].開始日時)) < Date.parse(tempstartUTC)) {
                    tempstart = astr[i].開始時間;
                    tempstartUTC = astr[i].開始日時;
                }
                //終了時間
                if (Date.parse(new Date(astr[i].終了日時)) > Date.parse(tempendingUTC)) {
                    tempending = astr[i].終了時間;
                    tempendingUTC = astr[i].終了日時;
                }
            }
        }
    }

    recd.出勤 = tempstart;
    recd.退勤 = tempending;
    //休憩時間処理（出勤時間＞12:00若しくは退勤時間＜13:00休憩時間計算しない）
    if (Date.parse(new Date(yr, mon, k + 1, 13, 0, 0)) > Date.parse(tempendingUTC) || Date.parse(new Date(yr, mon, k + 1, 12, 0, 0)) < Date.parse(tempstartUTC)) {
        recd.休憩時間 = ""
    } else {
        recd.休憩時間 = "1:00:00"
    }
    if (owner == "休日") {
        recd.出勤 = "";
        recd.退勤 = "";
    }
    recds[k] = recd;
}
//配列のエレメントのユニーク化
function unique(array) {
    var arr = [array[0]]; //Output
    //Index２から
    for (var i = 1; i < array.length; i++) {
        if (array.indexOf(array[i]) == i) arr.push(array[i]);
    }
    return arr;
}

// 日付比較　格式yyyy-mm-dd
function datecompare(a, b) {
    var starttime = new Date(new Date(a).getFullYear(), new Date(a).getMonth(), new Date(a).getDate());
    var starttimes = starttime.getTime();
    var endTime = new Date(new Date(b).getFullYear(), new Date(b).getMonth(), new Date(b).getDate());
    var endTimes = endTime.getTime();
    // 比较結果boolean
    return endTimes >= starttimes;
}


msg.payload = recds;
return msg;
