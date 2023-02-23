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
                        $('#view_data_department_wise').append(`<td>${item.EmployeeName}</td><td>${item.SectionName}</td><td>${item.CompanyName}</td><td>${item.DepartmentName}</td><td>${item.GradePoint}</td><td>${item.UnitPrice}</td>`);                
                        if(item.forecasts.length>0)   {
                            for(var i=0;i<item.forecasts.length;i++){
                                $('#view_data_department_wise').append(`<td>${item.forecasts[i].Points}</td>`);                                                    
                            }
                        }
                        $('#view_data_department_wise').append(`</tr>`);                
                    });                    
                    
                    

                    
                });
            }
            LoaderHide();
        },
        error: function () {           
        }
    });   
}
$('#add_name_search_button').on('click', function () {
    var departmentId = $('#view_export_department').val();
    if (departmentId == undefined || departmentId == null || departmentId == "") {
        alert("Select Department");
        return false;
    }
    LoaderShow();
    GetExportDataForView(departmentId);            
});

function LoaderShow() {
    $("#department_wise_list").css("display", "none");
    $("#loading").css("display", "block");
}
function LoaderHide() {
    $("#department_wise_list").css("display", "block");
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
$('#department').on('change', function () {
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