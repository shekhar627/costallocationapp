function onDepartmentInactiveClick() {
    let departmentIds = GetCheckedIds("department_list_tbody");    
    var apiurl = '/api/utilities/DepartmentCount?departmentIds=' + departmentIds;
    $.ajax({
        url: apiurl,
        type: 'Get',
        dataType: 'json',
        success: function (data) {
            $('.department_count').empty();
            $.each(data, function (key, item) {
                $('.department_count').append(`<li class='text-info'>${item}</li>`);
            });
        },
        error: function (data) {
        }
    });  
}

$(document).ready(function () {    
    GetDepartments();
    // LoadSectionDropdown();
    
    $('#department_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("department_list_tbody");
        id = id.slice(0, -1);

        $.ajax({
            url: '/api/departments?departmentIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetDepartments();
            },
            error: function (data) {
                ToastMessageFailed(data);

            }
        });

        $('#inactive_department_modal').modal('toggle');

    });

    $('#department_inactive_btn').on('click', function (event) {

        let id = GetCheckedIds("department_list_tbody");        
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });
});

//insert department
function InsertDepartment() {
    var apiurl = "/api/Departments/";
    let departmentName = $("#department_name").val().trim();
    // let sectionId = $("#section_list").val().trim();

    let isValidRequest = true;

    if (departmentName == "") {
        $(".department_name_err").show();
        isValidRequest = false;
    } else {
        $(".department_name_err").hide();
    }
    // if (sectionId == "" || sectionId < 0) {
    //     $("#section_ist_error").show();
    //     isValidRequest = false;
    // } else {
    //     $("#section_ist_error").hide();
    // }

    if (isValidRequest) {
        var data = {
            DepartmentName: departmentName
            // ,
            // SectionId: sectionId
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                $("#department_name").val('');
                $("#section_list").val('');

                ToastMessageSuccess(data);
                $('#allocation-name').val('');
                GetDepartments();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//get section list
function LoadSectionDropdown(){
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_list').empty();
            $('#section_list').append(`<option value='-1'>Select Section</option>`)
            $.each(data, function (key, item) {
                $('#section_list').append(`<option value='${item.Id}'>${item.SectionName}</option>`)
            });
        });
}
//get department list
function GetDepartments(){
    $.getJSON('/api/departments/')
    .done(function (data) {
        $('#department_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#department_list_tbody').append(`<tr><td><input type="checkbox" class="department_list_chk" data-id='${item.Id}' /></td><td>${item.DepartmentName}</td></tr>`);
        });
    });
}