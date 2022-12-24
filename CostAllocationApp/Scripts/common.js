//---------------------------master setting: get delete id list ------------------------------//
function GetCheckedIds(department_list_id) {
    var id = '';
    var sectionIds = $("#" + department_list_id + " tr input[type='checkbox']:checked").map(function () {
        return $(this).data('id')
    }).get();

    $.each(sectionIds, (index, data) => {
        id += data + ",";
    });
    return id;
}
//success toast message
function ToastMessageSuccess(response_data) {
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

//failed toast message
function ToastMessageFailed(response_data){
    Command: toastr["warning"](response_data.responseText, "Warning")

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
