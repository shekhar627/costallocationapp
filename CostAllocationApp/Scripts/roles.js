$(document).ready(function () {
    GetRoleList();
    
    $('#confirm_delete').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("incharge_list_tbody");

        id = id.slice(0, -1);
        console.log(id);

        $.ajax({
            url: '/api/Roles?roleIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);                
                GetRoleList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });

        $('#inactive_incharge_modal').modal('toggle');

    });

    $('#delete_button').on('click', function (event) {

        let id = GetCheckedIds("incharge_list_tbody");
        if (id == "") {
            alert("Please check first to delete.");
            return false;
        }
    });
});

function GetRoleList(){
    $.getJSON('/api/Roles/')
    .done(function (data) {
        $('#incharge_list_tbody').empty();
        $.each(data, function (key, item) {
            // Add a list item for the product.
            $('#incharge_list_tbody').append(`<tr><td><input type="checkbox" class="in_charge_list_chk" data-id='${item.Id}' /></td><td>${item.RoleName}</td></tr>`);
        });
    });
}