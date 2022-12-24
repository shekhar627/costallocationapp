$(document).ready(function () {
    //------------------Section Master----------------------//
    //show sections on page load
    GetSectionList();

    //check if section checked for inactive
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
                Command: toastr["warning"](data.responseText, "Warning")

                toastr.options = {
                    "closeButton": false,
                    "debug": false,
                    "newestOnTop": false,
                    "progressBar": false,
                    "positionClass": "toast-top-right",
                    "preventDuplicates": false,
                    "onclick": null,
                    "showDuration": "3000",
                    "hideDuration": "1000",
                    "timeOut": "5000",
                    "extendedTimeOut": "1000",
                    "showEasing": "swing",
                    "hideEasing": "linear",
                    "showMethod": "fadeIn",
                    "hideMethod": "fadeOut"
                }

            }
        });

        $('#delete_section').modal('toggle');

    });

    //success toast message
    function ToastMessageSuccess(response_data){
        Command: toastr["success"](response_data, "Success")
        toastr.options = {
            "closeButton": false,
            "debug": false,
            "newestOnTop": false,
            "progressBar": false,
            "positionClass": "toast-top-right",
            "preventDuplicates": false,
            "onclick": null,
            "showDuration": "3000",
            "hideDuration": "1000",
            "timeOut": "5000",
            "extendedTimeOut": "1000",
            "showEasing": "swing",
            "hideEasing": "linear",
            "showMethod": "fadeIn",
            "hideMethod": "fadeOut"
        }
    }

    //Get section list
    function GetSectionList(){
        $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_list_tbody').empty();
            $.each(data, function (key, item) {
                $('#section_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" data-id='${item.Id}' /></td><td>${item.SectionName}</td></tr>`);
            });
        });
    }
});