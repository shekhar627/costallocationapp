$(document).ready(function () {
    GetCompanyList();    

    $('#company_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("company_list_tbody");
        id = id.slice(0, -1);
        var companyWarningTxt = $("#company_warning_text").val();
        alert(companyWarningTxt);

        $.ajax({
            url: '/api/Companies?companyIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetCompanyList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });

        $('#inactive_company_modal').modal('toggle');

    });

    $('#company_inactive_btn').on('click', function (event) {
        let id = GetCheckedIds("company_list_tbody");
        IsCompanyAssigned(id);
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });

});
function IsCompanyAssigned(companyIds) {
    var returnVal = "";
    var apiurl = '/api/utilities/CompanyCount?companyIds=' + companyIds;
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
            $("#company_warning_text").val(returnVal);
        },
        error: function (data) {
            $("#company_warning_text").val(returnVal);
        }
    });
}

function GetCompanyList(){
    $.getJSON('/api/Companies/')
    .done(function (data) {
        $('#company_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#company_list_tbody').append(`<tr><td><input type="checkbox" class="company_list_chk" data-id='${item.Id}' /></td><td>${item.CompanyName}</td></tr>`);
        });
    });
}
function InsertCompanies() {
    var apiurl = "/api/Companies/";
    let companyName = $("#companyName").val().trim();
    if (companyName == "") {
        $(".company_name_err").show();
        return false;
    }
    else {
        $(".company_name_err").hide();
        var data = {
            CompanyName: companyName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                $("#companyName").val('');
                ToastMessageSuccess(data);
                GetCompanyList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
    }
}