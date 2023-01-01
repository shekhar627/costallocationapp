function onExplanationInactiveClick() {
    let explanationIds = GetCheckedIds("explanations_list_tbody");
    var apiurl = '/api/utilities/ExplanationCount?roleIds=' + explanationIds;
    $.ajax({
        url: apiurl,
        type: 'Get',
        dataType: 'json',
        success: function (data) {
            $('.explanation_count').empty();
            $.each(data, function (key, item) {
                $('.explanation_count').append(`<li class='text-info'>${item}</li>`);
            });
        },
        error: function (data) {
        }
    });
}

$(document).ready(function () {
    GetExplanationList();
       
    $('#explanations_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("explanations_list_tbody");
        id = id.slice(0, -1);
        
        $.ajax({
            url: '/api/Explanations?explanationIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetExplanationList();               
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
        $('#inactive_explanation_modal').modal('toggle');
    });

    $('#explanations_inactive_btn').on('click', function (event) {
        let id = GetCheckedIds("explanations_list_tbody");        
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });

});

function GetExplanationList(){
    $.getJSON('/api/Explanations/')
    .done(function (data) {
        $('#explanations_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#explanations_list_tbody').append(`<tr><td><input type="checkbox" class="explanation_list_chk" data-id='${item.Id}' /></td><td>${item.ExplanationName}</td></tr>`);
        });
    });
} 

function InsertExplanations() {
    var apiurl = "/api/Explanations/";
    let explanationName = $("#explanation_name").val().trim();
    if (explanationName == "") {
        $(".explanations_name_err").show();
        return false;
    }
    else {
        $(".explanations_name_err").hide();
        var data = {
            ExplanationName: explanationName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                $("#explanation_name").val('');
                ToastMessageSuccess(data);
                GetExplanationList();     
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
    }
}