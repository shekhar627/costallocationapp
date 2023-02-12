$(document).ready(function () {

    $.getJSON('/api/Grades/')
        .done(function (data) {
            $('#cm_grade').empty();
            $('#cm_grade').append('<option value="">Select Grade</option>');
            $.each(data, function (key, item) {
                $('#cm_grade').append(`<option value='${item.Id}'>${item.GradeName}</option>`);
            });
        });
});

//common master insert
function InsertCommonMaster() {
    var apiurl = "/api/commonmasters/";
    let gradeId = $("#cm_grade").val();
    let sir = $('#salary_increase_rate').val();
    let owft = $('#over_work_fixed_time').val();
    let brr = $('#bonus_reserve_ratio').val();
    let brc = $('#bonus_reserve_constant').val();
    let wcr = $('#welfare_cost_ratio').val();

    if (gradeId == "" || gradeId== undefined || gradeId==null) {
        $(".cm_grade_name_err").show();
        return false;
    } 
    else {
        $(".cm_grade_name_err").hide();
    }
    if (sir < 0) {
        $(".cm_sir_name_err").show();
        return false;
    }
    else {
        $(".cm_sir_name_err").hide();
    }
    if (owft < 0) {
        $(".cm_owft_name_err").show();
        return false;
    }
    else {
        $(".cm_owft_name_err").hide();
    }
    if (brr < 0) {
        $(".cm_brr_name_err").show();
        return false;
    }
    else {
        $(".cm_brr_name_err").hide();
    }
    if (brc < 0) {
        $(".cm_brc_name_err").show();
        return false;
    }
    else {
        $(".cm_brc_name_err").hide();
    }
    if (wcr < 0) {
        $(".cm_wcr_name_err").show();
        return false;
    }
    else {
        $(".cm_wcr_name_err").hide();
    }

    var data = {
        GradeId: gradeId,
        SalaryIncreaseRate: sir,
        OverWorkFixedTime: owft,
        BonusReserveRatio: brr,
        BonusReserveConstant: brc,
        WelfareCostRatio: wcr
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                ToastMessageSuccess(data);

                $('#salary_increase_rate').val('');
                $('#over_work_fixed_time').val('');
                $('#bonus_reserve_ratio').val('');
                $('#bonus_reserve_constant').val('');
                $('#welfare_cost_ratio').val('');
                GetCommonMasters();
            },
            error: function (data) {
                alert( data.responseJSON.Message);
            }
        });
}
//Get grade list
function GetCommonMasters() {
    $.getJSON('/api/CommonMasters/')
        .done(function (data) {
            $('#commonmaster_list_tbody').empty();
            $.each(data, function (key, item) {            
                //$('#grade_list_tbody').append(`<tr><td><input type="checkbox" class="grade_list_chk" onclick="GetCheckedIds(${item.Id});" data-id='${item.Id}' /></td><td>${item.GradeName}</td></tr>`);                
                $('#commonmaster_list_tbody').append(`<tr><td>${item.GradeName}</td><td>${item.SalaryIncreaseRate}</td><td>${item.OverWorkFixedTime}</td><td>${item.BonusReserveRatio}</td><td>${item.BonusReserveConstant}</td><td>${item.WelfareCostRatio}</td></tr>`);                
            });
        });
}

$(document).ready(function () {
    GetCommonMasters();
});
