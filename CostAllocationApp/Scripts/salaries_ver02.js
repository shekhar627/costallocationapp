//function onSalaryInactiveClick() {
//    let salaryIds = GetCheckedIds("salary_list_tbody");
//    var apiurl = '/api/utilities/SalaryCount?salaryIds=' + salaryIds;
//    $.ajax({
//        url: apiurl,
//        type: 'Get',
//        dataType: 'json',
//        success: function (data) {
//            $('.salary_count').empty();
//            $.each(data, function (key, item) {
//                $('.salary_count').append(`<li class='text-info'>${item}</li>`);
//            });
//        },
//        error: function (data) {
//        }
//    });
//}

$(document).ready(function () {
    //GetSalaries();

    //$('#salary_inactive_confirm_btn').on('click', function (event) {
    //    event.preventDefault();
    //    let id = GetCheckedIds("salary_list_tbody");
    //    id = id.slice(0, -1);        
    //    $.ajax({
    //        url: '/api/Salaries?salaryIds=' + id,
    //        type: 'DELETE',
    //        success: function (data) {
    //            ToastMessageSuccess(data);
    //            GetSalaries();                
    //        },
    //        error: function (data) {
    //            ToastMessageFailed(data);
    //        }
    //    });

    //    $('#inactive_salary_modal').modal('toggle');
    //});

    //$('#salary_inactive_btn').on('click', function (event) {

    //    let id = GetCheckedIds("salary_list_tbody");
    //    if (id == "") {
    //        alert("Please check first to delete.");
    //        return false;
    //    }
    //});

    


    $.getJSON('/api/departments/')
        .done(function (data) {
            $('#salary_department_input').empty();
            $('#salary_department_input').append(`<option value=''>Select Department</option>`);
            $.each(data, function (key, item) {
                $('#salary_department_input').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        });
    $.getJSON('/api/UnitPriceTypes/')
        .done(function (data) {
            $('#salary_type_input').empty();
            $('#salary_type_input').append(`<option value=''>Select Unit Price Type</option>`);
            $.each(data, function (key, item) {
                if (item.Id == 1 || item.Id == 4) {
                    $('#salary_type_input').append(`<option value='${item.Id}'>${item.TypeName}</option>`);
                }
            });
        });
    $.getJSON('/api/Grades/')
        .done(function (data) {
            $('#salary_grade_input').empty();
            $('#salary_grade_input').append(`<option value=''>Select Grade</option>`);
            $.each(data, function (key, item) {
                $('#salary_grade_input').append(`<option value='${item.Id}'>${item.GradeName}</option>`);
            });
        });




});
//function GetSalaries(){
//    $.getJSON('/api/Salaries/')
//    .done(function (data) {
//        $('#salary_list_tbody').empty();
//        $.each(data, function (key, item) {
//            let tempHighVal = "";
//            let tempLowVal = "";
//            if (item.SalaryLowPointWithComma == '' || item.SalaryLowPointWithComma == null) {
//                tempLowVal = "-"
//            } else {
//                tempLowVal = item.SalaryLowPointWithComma
//            }
//            if (item.SalaryHighPointWithComma == '' || item.SalaryHighPointWithComma == null) {
//                tempHighVal = "-"
//            } else {
//                tempHighVal = item.SalaryHighPointWithComma
//            }
//            $('#salary_list_tbody').append(`<tr><td><input type="checkbox" class="salary_list_chk" data-id='${item.Id}' /></td><td>${tempLowVal} ～ ${tempHighVal}</td><td>${item.SalaryGrade}</td></tr>`);
//        });
//    });
//}   

var temporaryDataArray = [];

function CreateGradeSalaryType() {
    let year = $('#salary_year_input').val();
    let departmentId = $('#salary_department_input').val();
    let departmentName = $('#salary_department_input option:selected').text();
    let salaryTypeId = $('#salary_type_input').val();
    var salaryTypeName = $('#salary_type_input option:selected').text();
    let gradeId = $('#salary_grade_input').val();
    let gradeName = $('#salary_grade_input option:selected').text();
    let beginning = $('#salary_amountlow_input').val();
    let revision = $('#salary_amounthigh_input').val();

    if (year == undefined || year == null || year == "") {
        $(".salary_year_err").show();
        return false;
    }
    else {
        $(".salary_year_err").hide();
    }

    if (departmentId == undefined || departmentId == null || departmentId == "") {
        $(".salary_department_name_err").show();
        return false;
    }
    else {
        $(".salary_department_name_err").hide();
    }

    if (salaryTypeId == undefined || salaryTypeId == null || salaryTypeId == "") {
        $(".salary_salaryType_err").show();
        return false;
    }
    else {
        $(".salary_type_input").hide();
    }

    if (gradeId == undefined || gradeId == null || gradeId == "") {
        $(".salary_gradeId_name_err").show();
        return false;
    }
    else {
        $(".salary_gradeId_name_err").hide();
    }

    if (beginning<0) {
        $(".salary_beginning_name_err").show();
        return false;
    }
    else {
        $(".salary_beginning_name_err").hide();
    }
    if (revision < 0) {
        $(".salary_revision_name_err").show();
        return false;
    }
    else {
        $(".salary_revision_name_err").hide();
    }

    var data = {
        GradeId: gradeId,
        GradeName: gradeName,
        GradeLowPoints: beginning,
        GradeHighPoints: revision,
        DepartmentId: departmentId,
        DepartmentName: departmentName,
        Year: year,
        SalaryTypeId: salaryTypeId,
        SalaryTypeName: salaryTypeName
    };
    temporaryDataArray.push(data);

    $('#temp_table tbody').empty();
    let count = 1;
    $.each(temporaryDataArray, function (index, item) {
        $('#temp_table tbody').append(`<tr><td>${count}</td><td>${item.Year}</td><td>${item.SalaryTypeName}</td><td>${item.DepartmentName}</td><td>${item.GradeName}</td><td>${item.GradeLowPoints}-${item.GradeHighPoints}</td><td><a href='javascript:void(0);' class='remove_item' onclick="onRemoveClick(this,${index});"><i class="fa fa-trash-o" aria-hidden="true"></i></a></td></tr>`);
        count++;
    });

    //$.ajax({
    //    url: apiurl,
    //    type: 'POST',
    //    dataType: 'json',
    //    data: data,
    //    success: function (data) {

    //        ToastMessageSuccess(data);
    //        GetAllSalaryTypes();
    //    },
    //    error: function (data) {
    //        //ToastMessageFailed(data);
    //    }
    //});
    
}


function confirmToSave() {
    if (temporaryDataArray.length > 0) {
        var apiurl = "/api/Salaries/";
        var confirmResult = confirm("Do you want to save?");
        let count = 0;
        let failedCount = 0;
        if (confirmResult == true) {
            $.each(temporaryDataArray, (index, item) => {
                $.ajax({
                    url: apiurl,
                    type: 'POST',
                    async: false,
                    dataType: 'json',
                    data: { GradeId: item.GradeId, GradeLowPoints: item.GradeLowPoints, GradeHighPoints: item.GradeHighPoints, DepartmentId: item.DepartmentId, Year: item.Year, SalaryTypeId: item.SalaryTypeId },
                    success: (data) => {
                        count++;
                        console.log(count);
                    },
                    error: (data) => {
                        failedCount++;
                    }
                });
            });

            temporaryDataArray = [];
            $('#temp_table tbody').empty();
            ToastMessageSuccess(count + " Data Saved Successfully!!!");
            ToastMessageFailed({ responseText: failedCount + " Data Cound Not Saved!!!" });
        }
    }
    else {
        ToastMessageFailed({ responseText: "No data found to saved!!!" });
    }

}

function clearData() {
    temporaryDataArray = [];
    $('#temp_table tbody').empty();
}

function onRemoveClick(element,arrayIndex) {
    
    var closestRow = $(element).closest('tr');
    var confirmResult = confirm('Do you want to remove?');
    if (confirmResult == true) {
        //removing element from table row
        closestRow.remove();
        // removing element from array
        temporaryDataArray.splice(arrayIndex, 1);
        // re-print the table
        $('#temp_table tbody').empty();
        let count = 1;
        $.each(temporaryDataArray, function (index, item) {
            $('#temp_table tbody').append(`<tr><td>${count}</td><td>${item.Year}</td><td>${item.SalaryTypeName}</td><td>${item.DepartmentName}</td><td>${item.GradeName}</td><td>${item.GradeLowPoints}-${item.GradeHighPoints}</td><td><a href='javascript:void(0);' class='remove_item' onclick="onRemoveClick(this,${index});"><i class="fa fa-trash-o" aria-hidden="true"></i></a></td></tr>`);
            count++;
        });
    }
}

//Get grade list
function GetAllSalaryTypes() {
    $.getJSON('/api/Salaries/')
        .done(function (data) {
            $('#GradeSalaryType_list_tbody').empty();
            $.each(data, function (key, item) {            
                //$('#grade_list_tbody').append(`<tr><td><input type="checkbox" class="grade_list_chk" onclick="GetCheckedIds(${item.Id});" data-id='${item.Id}' /></td><td>${item.GradeName}</td></tr>`);                
                $('#GradeSalaryType_list_tbody').append(`<tr><td>${item.GradeName}</td><td>${item.GradeLowPoints}</td><td>${item.GradeHighPoints}</td><td>${item.DepartmentName}</td><td>${item.SalaryTypeName}</td><td>${item.Year}</td></tr>`);                
            });
        });
}

$(document).ready(function () {
    GetAllSalaryTypes();
});

function ShowGradeTable() {
    $.getJSON('/api/utilities/GetSalaryMasterList/')
        .done(function (data) {
            $('#salary_master_list_tbl').empty();
            // <tr>
            //     <td rowspan="3">給与単価 (Salary Allowance Regular)</td>
            //     <td>G10</td>
            //     <td>1,250,000 </td>
            //     <td>1,250,000 </td>
            //     <td>1,250,000 </td>
            //     <td>1,250,000 </td>
            // </tr>
            var prevSalaryTypeId = 0;
            var tempCount = 0;
            console.log(data);
            $.each(data, function (key, item) {
                if (item.GradeSalaryTypes.length > 0) {
                    $('#salary_master_list_tbl').append(`<tr>`);
                    if (prevSalaryTypeId != item.UnitPriceType.Id) {
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.UnitPriceType.TypeName}</td>`);
                        prevSalaryTypeId = item.UnitPriceType.Id;
                    } else {
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'></td>`);
                    }
                    var isDevelopmentOk = false;
                    var isPlanningOk = false;

                    try {
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeName}</td>`);
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeLowWithCommaSeperate}</td>`);
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeHighWithCommaSeperate}</td>`);
                        isPlanningOk = true;
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`);
                        $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`);
                        isPlanningOk = true;
                    } catch (error) {
                        if (isPlanningOk) {
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`);
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`);
                        } else {
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`);
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`);
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`);
                            $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`);
                        }
                    }
                    // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeName}</td>`); 
                    // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeLowWithCommaSeperate}</td>`); 
                    // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeHighWithCommaSeperate}</td>`); 
                    // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`); 
                    // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`); 

                    $('#salary_master_list_tbl').append(`</tr>`);
                }
            });
        });

    $("#salary_master_list").css("display", "block");
}


//function ShowGradeTable(){
//    $.getJSON('/api/utilities/GetSalaryMasterList/')
//    .done(function (data) {
//    $('#salary_master_list_tbl').empty();
//    // <tr>
//    //     <td rowspan="3">給与単価 (Salary Allowance Regular)</td>
//    //     <td>G10</td>
//    //     <td>1,250,000 </td>
//    //     <td>1,250,000 </td>
//    //     <td>1,250,000 </td>
//    //     <td>1,250,000 </td>
//    // </tr>
//    var prevSalaryTypeId = 0;
//    var tempCount = 0;
//    $.each(data, function(key, item) {
//        if(item.GradeSalaryTypes.length > 0){
//            $('#salary_master_list_tbl').append(`<tr>`);                
//            if (prevSalaryTypeId != item.SalaryType.Id){                        
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.SalaryType.SalaryTypeName}</td>`);  
//                prevSalaryTypeId = item.SalaryType.Id;
//            }else{
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'></td>`);  
//            }
//            var isDevelopmentOk = false;
//            var isPlanningOk = false;

//            try { 
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeName}</td>`); 
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeLowWithCommaSeperate}</td>`); 
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[0].GradeHighWithCommaSeperate}</td>`); 
//                isPlanningOk = true;
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`); 
//                $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`); 
//                isPlanningOk = true;
//              } catch (error) {     
//                if(isPlanningOk)       {
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`); 
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`); 
//                }   else{
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`); 
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>0</td>`); 
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`); 
//                    $('#salary_master_list_tbl').append(`<td class='gradelist_td'>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`); 
//                }             
//              }
//            // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeName}</td>`); 
//            // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeLowWithCommaSeperate}</td>`); 
//            // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[0].GradeHighWithCommaSeperate}</td>`); 
//            // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[1].GradeLowWithCommaSeperate}</td>`); 
//            // $('#salary_master_list_tbl').append(`<td>${item.GradeSalaryTypes[1].GradeHighWithCommaSeperate}</td>`); 

//            $('#salary_master_list_tbl').append(`</tr>`);                
//        }
//    });  
//});

//$("#salary_master_list").css("display", "block");
//}

$('#add_name_search_button').on('click', function (event) {
    
    var selectYear = $('#salary_year_list').val();
    if(selectYear >0){
        ShowGradeTable();
       
    }else{        
        $("#salary_master_list").css("display", "none");
        $( "#salary_year_list" ).focus();        
        alert("please select year!");
    }
});
function SaveSalaryType() {
    let year = $('#salary_year_input').val();
    console.log("year: " + year)

    let salaryTypeId = $('#salary_type_input').val();
    console.log("salaryTypeId: " + salaryTypeId)
    var salaryTypeName = $('#salary_type_input option:selected').text();

    let departmentId = $('#salary_department_input').val();
    console.log("departmentId: " + departmentId)
    let departmentName = $('#salary_department_input option:selected').text();

    let gradeId = $('#salary_grade_input').val();
    console.log("gradeId: " + gradeId)
    let gradeName = $('#salary_grade_input option:selected').text();

    let beginning = $('#salary_amountlow_input').val();
    console.log("beginning: " + beginning)

    let revision = $('#salary_amounthigh_input').val();
    console.log("revision: " + revision)

    if (year == undefined || year == null || year == "") {
        alert("select year");
        return false;
    }
    if (salaryTypeId == undefined || salaryTypeId == null || salaryTypeId == "") {
        alert("select salary type");
        return false;
    }
    if (departmentId == undefined || departmentId == null || departmentId == "") {
        alert("select department");
        return false;
    }

    if (gradeId == undefined || gradeId == null || gradeId == "") {
        alert("select grade");
        return false;
    }

    if (beginning < 0) {
        alert("input low points");
        return false;
    }
    if (revision < 0) {
        alert("input high points");
        return false;
    }

    var data = {
        GradeId: gradeId,
        GradeName: gradeName,
        GradeLowPoints: beginning,
        GradeHighPoints: revision,
        DepartmentId: departmentId,
        DepartmentName: departmentName,
        Year: year,
        UnitPriceTypeId: salaryTypeId,
        SalaryTypeName: salaryTypeName
    };
    var apiurl = "/api/Salaries/";
    let count = 0;
    let failedCount = 0;
    $.ajax({
        url: apiurl,
        type: 'POST',
        async: false,
        dataType: 'json',
        data: data,
        //data: { GradeId: item.GradeId, GradeLowPoints: item.GradeLowPoints, GradeHighPoints: item.GradeHighPoints, DepartmentId: item.DepartmentId, Year: item.Year, SalaryTypeId: item.SalaryTypeId },
        success: function (data) {
            alert("success");
            ShowGradeTable();
        },
        error: function () {
            alert("failed");
        }
    });

    // if (temporaryDataArray.length > 0) {


    //     // temporaryDataArray = [];
    //     // $('#temp_table tbody').empty();
    //     // ToastMessageSuccess(count + " Data Saved Successfully!!!");
    //     // ToastMessageFailed({ responseText: failedCount + " Data Cound Not Saved!!!" });
    // }
    // else {
    //     ToastMessageFailed({ responseText: "No data found to saved!!!" });
    // } 
}

//function SaveSalaryType(){
//    let year = $('#salary_year_input').val();
//    console.log("year: "+year)

//    let salaryTypeId = $('#salary_type_input').val();
//    console.log("salaryTypeId: "+salaryTypeId)    
//    var salaryTypeName = $('#salary_type_input option:selected').text();

//    let departmentId = $('#salary_department_input').val();
//    console.log("departmentId: "+departmentId)    
//    let departmentName = $('#salary_department_input option:selected').text();    

//    let gradeId = $('#salary_grade_input').val();
//    console.log("gradeId: "+gradeId)
//    let gradeName = $('#salary_grade_input option:selected').text();

//    let beginning = $('#salary_amountlow_input').val();
//    console.log("beginning: "+beginning)

//    let revision = $('#salary_amounthigh_input').val();
//    console.log("revision: "+revision)

//    if (year == undefined || year == null || year == "") {
//        alert("select year");
//        return false;
//    }    
//    if (salaryTypeId == undefined || salaryTypeId == null || salaryTypeId == "") {
//        alert("select salary type");
//        return false;
//    }
//    if (departmentId == undefined || departmentId == null || departmentId == "") {
//        alert("select department");
//        return false;
//    }
    
//    if (gradeId == undefined || gradeId == null || gradeId == "") {
//        alert("select grade");
//        return false;
//    }

//    if (beginning<0) {
//        alert("input low points");
//        return false;
//    }
//    if (revision < 0) {
//        alert("input high points");
//        return false;
//    }

//    var data = {
//        GradeId: gradeId,
//        GradeName: gradeName,
//        GradeLowPoints: beginning,
//        GradeHighPoints: revision,
//        DepartmentId: departmentId,
//        DepartmentName: departmentName,
//        Year: year,
//        SalaryTypeId: salaryTypeId,
//        SalaryTypeName: salaryTypeName
//    };
//    var apiurl = "/api/Salaries/";
//    let count = 0;
//    let failedCount = 0;
//    $.ajax({
//        url: apiurl,
//        type: 'POST',
//        async: false,
//        dataType: 'json',
//        data:data,
//        //data: { GradeId: item.GradeId, GradeLowPoints: item.GradeLowPoints, GradeHighPoints: item.GradeHighPoints, DepartmentId: item.DepartmentId, Year: item.Year, SalaryTypeId: item.SalaryTypeId },
//        success: function(data){
//            alert("success");
//            ShowGradeTable();
//        },
//        error: function() {
//            alert("failed");
//        }
//    });

//    // if (temporaryDataArray.length > 0) {
        

//    //     // temporaryDataArray = [];
//    //     // $('#temp_table tbody').empty();
//    //     // ToastMessageSuccess(count + " Data Saved Successfully!!!");
//    //     // ToastMessageFailed({ responseText: failedCount + " Data Cound Not Saved!!!" });
//    // }
//    // else {
//    //     ToastMessageFailed({ responseText: "No data found to saved!!!" });
//    // } 
//}

//scroll purpose
// Change the selector if needed
var $table = $('table'),
    $bodyCells = $table.find('tbody tr:first').children(),
    colWidth;

// Get the tbody columns width array
colWidth = $bodyCells.map(function() {
    return $(this).width();
}).get();

// Set the width of thead columns
$table.find('thead tr').children().each(function(i, v) {
    $(v).width(colWidth[i]);
});