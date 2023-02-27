function GetExportDataForView(departmentId){    
    $.ajax({
        url: `/exports/ViewDataByDepartment/`,
        type: 'GET',
        dataType: 'json',        
        //data:'gradeId='+gradeId+"&departmentId="+departmentId,
        data:"departmentId="+departmentId,
        success: function (data) {
            let dataCount = data.length;
            if(dataCount >0){
                $.each(data, function (key, item) {
                    var employeeName = item.EmployeeName;
                    var sectionName = item.SectionName;
                    var companynName = item.CompanyName;
                    var departmentName = item.DepartmentName;
                    var gradePointName = item.GradePoint;
                    var unitPrice = item.UnitPrice;
                    
                    //var sheet.Cells["F" + count].Value = String.IsNullOrEmpty(item.UnitPrice) == true ? "" : Convert.ToDecimal(item.UnitPrice).ToString("N0");
                    $('#view_data_department_wise').empty();
                                   
                    $.each(data, function (key, item) {      
                        $('#view_data_department_wise').append(`<tr>`);                               
                        $('#view_data_department_wise').append(`<td style='text-align:left;'>${item.EmployeeName}</td><td style='text-align:center;'>${item.SectionName}</td><td>${item.CompanyName}</td><td>${item.DepartmentName}</td><td>${item.GradePoint}</td><td class='department_wise_list_unit_price'>${item.UnitPrice}</td>`);                
                        if(item.forecasts.length>0)   {
                            for(var i=0;i<item.forecasts.length;i++){
                                $('#view_data_department_wise').append(`<td>${item.forecasts[i].Points}</td>`);                                                    
                            }
                            for (var i = 0; i < item.forecasts.length; i++) {
                                $('#view_data_department_wise').append(`<td style='text-align:right;'>${item.forecasts[i].Total}</td>`);
                            }
                        }
                        $('#view_data_department_wise').append(`</tr>`);                
                    });                    
                    
                    

                    
                });
            }
            LoaderHide('department_wise_list');
        },
        error: function () {           
        }
    });   
}
function AllocationWiseDataView(departmentId,explanationId){    

    $.ajax({
        url: `/exports/ViewDataByAllocation/`,
        type: 'GET',
        dataType: 'json',        
        //data:'gradeId='+gradeId+"&departmentId="+departmentId,
        data:"departmentId="+departmentId+"&explanationId="+explanationId,
        success: function (data) {
            $("#allcoationlist_tbl").html(data.AllocationWiseExportList);
            // let dataCount = data.length;
            // if(dataCount >0){
            //     $.each(data, function (key, item) {
            //         var employeeName = item.EmployeeName;
            //         var sectionName = item.SectionName;
            //         var companynName = item.CompanyName;
            //         var departmentName = item.DepartmentName;
            //         var gradePointName = item.GradePoint;
            //         var unitPrice = item.UnitPrice;
                    
            //         //var sheet.Cells["F" + count].Value = String.IsNullOrEmpty(item.UnitPrice) == true ? "" : Convert.ToDecimal(item.UnitPrice).ToString("N0");
            //         $('#view_data_department_wise').empty();
                                   
            //         $.each(data, function (key, item) {      
            //             $('#view_data_department_wise').append(`<tr>`);                               
            //             $('#view_data_department_wise').append(`<td>${item.EmployeeName}</td><td>${item.SectionName}</td><td>${item.CompanyName}</td><td>${item.DepartmentName}</td><td>${item.GradePoint}</td><td>${item.UnitPrice}</td>`);                
            //             if(item.forecasts.length>0)   {
            //                 for(var i=0;i<item.forecasts.length;i++){
            //                     $('#view_data_department_wise').append(`<td>${item.forecasts[i].Points}</td>`);                                                    
            //                 }
            //             }
            //             $('#view_data_department_wise').append(`</tr>`);                
            //         });                    
                    
                    

                    
            //     });
            // }
            
            LoaderHide('alocation_wise_list');
        },
        error: function () {           
        }
    });   
}
$('#department_wise_search_btn').on('click', function () {
    var departmentId = $('#view_export_department').val();
    if (departmentId == undefined || departmentId == null || departmentId == "") {
        alert("Select Department");
        return false;
    }
    $("#alocation_wise_list").css("display", "none");
    LoaderShow('department_wise_list');
    GetExportDataForView(departmentId);              
});

$('#allocation_wise_search_btn').on('click', function () {
    var departmentId = $('#department_allocation_wise').val();
    var explanationId = $('#explanation').val();

    var isDepeartmentEmpty = false; 
    var isAllocationEmpty = false; 

    if (departmentId == undefined || departmentId == null || departmentId == "") {
        isDepeartmentEmpty = true;
    }

    if (explanationId == undefined || explanationId == null || explanationId == "") {
        isAllocationEmpty = true; 
    }
    if (isDepeartmentEmpty) {
        alert("Select Department");
        return false;
    } else if (!isDepeartmentEmpty && isAllocationEmpty) {
        $("#alocation_wise_list").css("display", "none");
        LoaderShow('department_wise_list');
        GetExportDataForView(departmentId);
    } else if (!isDepeartmentEmpty && !isAllocationEmpty) {
        $("#department_wise_list").css("display", "none");
        LoaderShow('alocation_wise_list');
        AllocationWiseDataView(departmentId, explanationId); 
    }          
});
function LoaderShow(tableId) {
    $("#"+tableId).css("display", "none");
    $("#loading").css("display", "block");
}
function LoaderHide(tableId) {
    $("#"+tableId).css("display", "block");
    $("#loading").css("display", "none");
}


$('#export').on('click', function (e) {
    var departmentValue = $('#department').val();
    var explanationValue = $('#explanation').val();
    if (departmentValue == undefined || departmentValue == null || departmentValue == "") {
        alert("Select Department");
        e.preventDefault();
    }
    if (explanationValue == undefined || explanationValue == null || explanationValue == "") {
        alert("Select Explanation");
        e.preventDefault();
    }
});
// $('#department').on('change', function () {
//     var departmentId = $(this).val();
//     $.getJSON('/api/Utilities/GetAllExplanationsByDepartmentId?departmentId=' + departmentId)
//         .done(function (data) {
//             $('#explanation').empty();
//             $('#explanation').append(`<option value=''>Select Allocation</option>`)
//             $.each(data, function (key, item) {
//                 $('#explanation').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`)
//             });
//         });
// });

$('#department_allocation_wise').on('change', function () {
    var departmentId = $(this).val();
    $.ajax({
        url: `/exports/GetAllExplanationsByDepartmentId/`,
        type: 'GET',
        dataType: 'json',        
        data:"departmentId="+departmentId,
        success: function (data) {
            $('#explanation').empty();
            $('#explanation').append(`<option value=''>select allocation</option>`)
            $.each(data, function (key, item) {
                $('#explanation').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`)
            });
        },
        error: function () {           
        }
    });   
});


function LoaderShow_AllData() {
    $("#btn_export_all_data").prop("disabled",true);
    $("#department_allocation_export").css("display", "none");
    $("#all_data_loader").css("display", "block");
}
function LoaderHide_AllData() { 
    setTimeout(function() {
        $("#department_allocation_export").css("display", "block");
        $("#btn_export_all_data").prop("disabled",false);
        $("#all_data_loader").css("display", "none");    

    }, 6000);

    // $("#all_data_loader").delay(3000).queue(function() {
    //     $("#department_allocation_export").css("display", "block");
    //     $("#all_data_loader").css("display", "none");        
    //     $("#btn_export_all_data").prop("disabled",false);
    //   });            
    //$("#all_data_loader").css("display", "none");
}

$('#btn_export_all_data').click(function(){
    LoaderShow_AllData();
    window.location.href='/Exports/ExportAllData';
    LoaderHide_AllData();
 })

 function ExportByDepartment(){
    var departmentId = $('#department_allocation_wise').val();
    $('#department').val(departmentId).attr("selected", "selected");
    $('#frmDepartmentWiseExport').submit();
 }
 function ExportByAllocation(){
    var departmentId = $('#department_allocation_wise').val();
    var explanationId = $('#explanation').val();
    $('#departmentId').val(departmentId);
    $('#explanationId').val(explanationId);

    $('#frmAllocationWiseExport').submit();
 }