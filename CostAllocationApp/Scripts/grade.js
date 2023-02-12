//function onSectionInactiveClick() {
//    let sectionIds = GetCheckedIds("section_list_tbody");
//    var apiurl = '/api/utilities/SectionCount?sectionIds=' + sectionIds;
//    $.ajax({
//        url: apiurl,
//        type: 'Get',
//        dataType: 'json',
//        success: function (data) {
//            $('.section_count').empty();
//            $.each(data, function (key, item) {
//                $('.section_count').append(`<li class='text-info'>${item}</li>`);
//            });
//        },
//        error: function (data) {
//        }
//    });
//}

//grade insert
function InsertGrade() {
    var apiurl = "/api/grades/";
    let gradeName = $("#grade-name").val().trim();
    if (gradeName == "") {
        $(".grade_name_err").show();
        return false;
    } else {
        $(".grade_name_err").hide();
        var data = {
            GradeName: gradeName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                ToastMessageSuccess(data);

                $('#grade-name').val('');
                GetGradeList();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//Get grade list
function GetGradeList() {
    $.getJSON('/api/grade/')
        .done(function (data) {
            $('#grade_list_tbody').empty();
            $.each(data, function (key, item) {
                $('#grade_list_tbody').append(`<tr><td><input type="checkbox" class="grade_list_chk" onclick="GetCheckedIds(${item.Id});" data-id='${item.Id}' /></td><td>${item.GradeName}</td></tr>`);
            });
        });
}