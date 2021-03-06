// Method to upload a valid excel file
var backup;
var edit;

function upload() {
    var files = document.getElementById('file_upload').files;
    if(files.length==0){
      alert("Please choose any file...");
      return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        //Here calling another method to read excel file into json
        excelFileToJSON(files[0]);
        
    }else{
        alert("Please select a valid excel file.");
    }
  }

//Method to read excel file and convert it into JSON 
function excelFileToJSON(file){
    try {
      var reader = new FileReader();
      reader.readAsBinaryString(file);
      reader.onload = function(e) {

          var data = e.target.result;
          var workbook = XLSX.read(data, {
              type : 'binary'
          });
          var result = {};
          var firstSheetName = workbook.SheetNames[0];
          //reading only first sheet data
          var jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
          alert(JSON.stringify(jsonData));
          //displaying the json result into HTML table
          displayJsonToHtmlTable(jsonData);
          backup = JSON.parse(JSON.stringify(jsonData));
          edit = JSON.parse(JSON.stringify(jsonData));
          }
      }catch(e){
          console.error(e);
      }
}  

//Method to display the data in HTML Table
function displayJsonToHtmlTable(jsonData){
    var table=document.getElementById("display_excel_data");
    var avg_element=document.getElementById("avg");
    var avg_numerator=0.0,avg_denominator=0.0,avg=0.0;
    if(jsonData.length>0){
        var htmlData='<tr><th>id_subject</th><th>Name_subject</th><th>TC</th><th>score_10</th><th>word</th><th>score_4</th><th>disable</th></tr>';
        for(var i=0;i<jsonData.length;i++){
            var row=jsonData[i];
            var score4 =convertScore4(row["score_10"]);
            avg_numerator += parseFloat(score4.number)*parseFloat(row["TC"]);
            avg_denominator += parseFloat(row["TC"]);
            htmlData+='<tr><td>'+row["id_subject"]+'</td><td>'+row["Name_subject"]
                  +'</td><td>'+row["TC"]+'</td>'+`<td><input onkeyup="editScore('${row["id_subject"]}')" class="form-control"  type="text" id="${row["id_subject"]}" value="${row["score_10"]}" />`+
                  '</td><td>'+score4.word+'</td><td>'+score4.number+'</td>'+
                  `<td><button class="btn btn-light" onclick="editScore('${row["id_subject"]}','1')"></button></td>`+
                  '</tr>'

                  ;
        }
        avg = (avg_numerator/avg_denominator)
        //htmlData += `<tr><td colspan="6"> t??nh trung b??nh t??ch l??y: ${avg.toFixed(2)} </td></tr>`;
        avg_element.innerHTML = `<tr><td colspan="6">trung b??nh t??ch l??y: ${avg.toFixed(2)} </td></tr>`;
        table.innerHTML=htmlData ;


    }else{
        table.innerHTML='There is no data in Excel';
    }
}
function convertScore4(score){
    var score4= {
        word:"N",
        number:0
    };
    if (score<4){
        score4.word ='F';
        score4.number = 0;
    }
    else if(4<=score && score<=4.9){
        score4.word ='D';
        score4.number = 1;
    }
    else if(5<=score && score<=5.4){
        score4.word ='D+';
        score4.number = 1.5;
    }
    else if(5.5<=score && score<=6.4){
        score4.word ='C';
        score4.number = 2;
    }  
    else if(6.5<=score && score<=6.9){
        score4.word ='C+';
        score4.number = 2.5;
    }
    else if(7<=score && score<=7.9){
        score4.word ='B';
        score4.number = 3;
    } 
    else if(8<=score && score<=8.9){
        score4.word ='B+';
        score4.number = 3.5;
    }
    else {
        score4.word ='A';
        score4.number = 4;
    }                    
    return score4
}
function back(){

    displayJsonToHtmlTable(backup);
}

function editScore(id,check){

    var temp = document.querySelector("#"+id)
    var index = edit.findIndex( x => x.id_subject === id );

    if (check == 1){
        var tc = backup[index].TC;
        if(edit[index].TC>0){
            edit[index].TC = 0;
        }
        else{
            edit[index].TC = tc;
        }
    }

    edit[index]["score_10"] = parseFloat(temp.value);
    
    if(temp.value.length == 3 || check==1){

        displayJsonToHtmlTable(edit);
    }

}





