$(document).ready(function () {
    //------------------Section Master----------------------//

    //show sections on page load
    GetSectionList();

    //is section checked
    $('#section_inactive_btn').on('click', function (event) {
        let id = GetCheckedIds("section_list_tbody");
        IsSectionAssigned(id);
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });
    //inactive section
    $('#section_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("section_list_tbody");

        var sectionWarningTxt = $("#section_warning_text").val();
        $("#section_warning").html(sectionWarningTxt);
        var tempVal = $("#section_warning").html();
        alert(tempVal)
       

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
//function GetCheckedIds(sectionId) {
//    var hidSectionIds = $("#section_chk_ids").val();
//    if (hidSectionIds == '') {
//        $("#section_chk_ids").val(hidSectionIds);
//    } else {

//    }
//    alert("test: " + sectionId);
//}

//Get Assgined Section Count
function IsSectionAssigned(sectionIds) {
    var returnVal = "";
    var apiurl = '/api/utilities/SectionCount?sectionIds=' + sectionIds;
    $.ajax({
        url: apiurl,
        type: 'Get',
        dataType: 'json',
        success: function (data) {
            $.each(data, function (key, item) {
                if (returnVal == "") {
                    returnVal = item;
                } else {
                    returnVal = returnVal + ". " + item;
                }
            });
            $("#section_warning_text").val(returnVal);
        },
        error: function (data) {
            $("#section_warning_text").val(returnVal);
        }
    });    
}

//sections insert
function InsertSection() {
    var apiurl = "/api/sections/";
    let sectionName = $("#section-name").val().trim();
    if (sectionName == "") {
        $(".section_name_err").show();
        return false;
    } else {
        $(".section_name_err").hide();
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
                $('#section_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" onclick="GetCheckedIds(${item.Id});" data-id='${item.Id}' /></td><td>${item.SectionName}</td></tr>`);
            });
        });
}