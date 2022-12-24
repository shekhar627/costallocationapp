﻿$(document).ready(function () {
    //------------------Section Master----------------------//

    //show sections on page load
    GetSectionList();

    //is section checked
    $('#section_inactive_btn').on('click', function (event) {
        let id = GetCheckedIds("section_list_tbody");
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });

    //inactive section
    $('#section_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("section_list_tbody");
        id = id.slice(0, -1);
        $.ajax({
            url: '/api/sections?sectionIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetSectionList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
        $('#delete_section').modal('toggle');
    });
});

//sections insert
function InsertSection() {
    var apiurl = "/api/sections/";
    let sectionName = $("#section-name").val().trim();
    if (sectionName == "") {
        $(".section_name_err").show();
        return false;
    }
    else {
        var data = {
            SectionName: sectionName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                ToastMessageSuccess(data);
                
                $('#section-name').val('');
                GetSectionList();                
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//Get section list
function GetSectionList() {
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_list_tbody').empty();
            $.each(data, function (key, item) {
                $('#section_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" data-id='${item.Id}' /></td><td>${item.SectionName}</td></tr>`);
            });
        });
}  