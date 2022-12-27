$(document).ready(function () {    
    GetInchargeList();

    $('#incharge_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("incharge_list_tbody");

        id = id.slice(0, -1);
        var inchargeWarningTxt = $("#incharge_warning_text").val();
        alert(inchargeWarningTxt);

        $.ajax({
            url: '/api/InCharges?inChargeIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetInchargeList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });

        $('#del_incharge_modal').modal('toggle');

    });

    $('#incharge_inactive_btn').on('click', function (event) {

        let id = GetCheckedIds("incharge_list_tbody");
        IsInchargeAssigned(id);
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        } else {

        }
    });
});

function IsInchargeAssigned(inChargeIds) {
    var returnVal = "";
    var apiurl = '/api/utilities/InChargeCount?inChargeIds=' + inChargeIds;
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
            $("#incharge_warning_text").val(returnVal);
        },
        error: function (data) {
            $("#incharge_warning_text").val(returnVal);
        }
    });
}

//insert department
function InsertInCharge() {
    var apiurl = "/api/incharges/";
    let in_charge_name = $("#in_charge_name").val().trim();
    if (in_charge_name == "") {
        $(".incharge_name_err").show();
        return false;
    }
    else {
        $(".incharge_name_err").hide();
        var data = {
            InChargeName: in_charge_name
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                $("#in_charge_name").val('');
                ToastMessageSuccess(data)
                $('#section-name').val('');

                GetInchargeList();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//get department list
function GetInchargeList(){
    $.getJSON('/api/InCharges/')
    .done(function (data) {
        $('#incharge_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.InChargeName}</td></tr>`);
        });
    });
}

//insert department
function InsertInCharge() {
    var apiurl = "/api/incharges/";
    let in_charge_name = $("#in_charge_name").val().trim();
    if (in_charge_name == "") {
        $(".incharge_name_err").show();
        return false;
    }
    else {
        $(".incharge_name_err").hide();
        var data = {
            InChargeName: in_charge_name
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                $("#in_charge_name").val('');
                ToastMessageSuccess(data)
                $('#section-name').val('');

                GetInchargeList();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//get department list
function GetInchargeList(){
    $.getJSON('/api/InCharges/')
    .done(function (data) {
        $('#incharge_list_tbody').empty();
        $.each(data, function (key, item) {
            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.InChargeName}</td></tr>`);
        });
    });
}