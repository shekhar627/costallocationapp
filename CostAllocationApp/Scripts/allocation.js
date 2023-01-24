function onAlllocationInactiveClick() {
    let allocationIds = GetCheckedIds("allocation_list_tbody");
    var apiurl = '/api/utilities/AllocationCount?allocationIds=' + allocationIds;
    $.ajax({
        url: apiurl,
        type: 'Get',
        dataType: 'json',
        success: function (data) {
            $('.allocation_count').empty();
            $.each(data, function (key, item) {
                $('.allocation_count').append(`<li class='text-info'>${item}</li>`);
            });
        },
        error: function (data) {
        }
    });  
}
$(document).ready(function () {
    //------------------Section Master----------------------//

    //show sections on page load
    GetAllocationList();

    //inactive section
    $('#allocation_inactive_confirm_btn').on('click', function (event) {
        event.preventDefault();
        let id = GetCheckedIds("allocation_list_tbody");

        var sectionWarningTxt = $("#allocation_warning_text").val();
        $("#allocation_warning").html(sectionWarningTxt);
        var tempVal = $("#allocation_warning").html();
        //alert(tempVal)
       

        id = id.slice(0, -1);
        $.ajax({
            url: '/api/allocation?allocationIds=' + id,
            type: 'DELETE',
            success: function (data) {
                ToastMessageSuccess(data);
                GetAllocationList();
            },
            error: function (data) {
                ToastMessageFailed(data);
            }
        });
        $('#delete_allocation').modal('toggle');
    });
});
$('#allocation_inactive_btn').on('click', function (event) {

    let id = GetCheckedIds("allocation_list_tbody");
    if (id == "") {
        alert("Please check first to delete.");
        return false;
    }
});

//Get Assgined Section Count
function IsSectionAssigned(allocationIds) {
    var returnVal = "";
    var apiurl = '/api/utilities/AllocationCount?allocationIds=' + allocationIds;
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
            $("#allocation_warning_text").val(returnVal);
        },
        error: function (data) {
            $("#allocation_warning_text").val(returnVal);
        }
    });    
}

//Allocation insert
function InsertAllocation() {
    var apiurl = "/api/allocation/";
    let allocationName = $("#allocation-name").val().trim();
    if (allocationName == "") {
        $(".allocation_name_err").show();
        return false;
    } else {
        $(".allocation_name_err").hide();
        var data = {
            AllocationName: allocationName
        };

        $.ajax({
            url: apiurl,
            type: 'POST',
            dataType: 'json',
            data: data,
            success: function (data) {
                $("#page_load_after_modal_close").val("yes");
                ToastMessageSuccess(data);

                $('#allocation-name').val('');
                GetAllocationList();
            },
            error: function (data) {
                alert(data.responseJSON.Message);
            }
        });
    }
}

//Get section list
function GetAllocationList() {
    $.getJSON('/api/allocation/')
        .done(function (data) {
            $('#allocation_list_tbody').empty();
            $.each(data, function (key, item) {
                $('#allocation_list_tbody').append(`<tr><td><input type="checkbox" class="section_list_chk" onclick="GetCheckedIds(${item.Id});" data-id='${item.Id}' /></td><td>${item.AllocationName}</td></tr>`);
            });
        });
}