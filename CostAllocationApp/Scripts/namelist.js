$(document).ready(function() {
    var count=1;
    $('#namelist thead tr:eq(1) th').each( function () {
        if(count == 1){
            var title = $(this).text();
            $(this).html( '<input type="text" placeholder="Search '+title+'" class="column_search" Id="name_search"/>' );
        }  
        count = count +1;      
    } );
    // DataTable
    // var table = $('#namelist').DataTable({
    //     orderCellsTop: true,
    //     searching: false,
    //     bLengthChange: false,
    //     orderCellsTop: true,
    //     columnDefs: [
    //         { width: 10, targets: 0 }
    //     ]
    // });

    // Apply the search
    // $( '#namelist thead'  ).on( 'keyup', ".column_search",function () {   
    //     table
    //         .column( $(this).parent().index() )
    //         .search( this.value )
    //         .draw();
    // } );


    // //section multi search
    $('#section_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select section',
    });
    //department multi search
    $('#dept_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select dept',
    });
    //incharge multi search
    $('#incharge_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select incharge',
    });
    //role multi search
    $('#role_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select role',
    });
    //explanation multi search
    $('#explanation_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select explanation',
    });
    //company multi search
    $('#company_multi_search').multiselect({
        includeSelectAllOption: true,
        enableFiltering: true,
        nonSelectedText: 'select company',
    });
    
    LoadNameListTableOnLoad();

    function LoadNameListTableOnLoad() {
        let employeeName = "";
        let sectionId = "";
        let departmentId = "";
        let inchargeId = "";
        let roleId = "";
        let explanationId = "";
        let companyId = "";
        let status = "";

        var url = '@Url.Action("SearchEmployee", "/api/utilities")'; // don't hard code your urls!
        $.getJSON(`/api/utilities/SearchEmployee`, {
                employeeName: employeeName,
                sectionId: sectionId,
                departmentId: departmentId,
                inchargeId: inchargeId,
                roleId: roleId,
                explanationId: explanationId,
                companyId: companyId,
                status: status
            })
            .done(function(data) {
                $('#employee_list_search_results').empty();

                //Name list with datatable
                NameList_DatatableLoad(data);

                let dataCount = data.length;
                if (dataCount > 0) {
                    let counter = 1;
                    let nameWithCodeAndRemarks = "";                    
                } else {
                    $('#employee_list_search_results').append("<tr><td colspan='8' style='text-align:center;font-size:20px;'>No Data found!</td></tr>");

                }

            });
        GetListDropdownValue();
    }
    
    function GetListDropdownValue() {
        $.getJSON('/api/sections/')
            .done(function(data) {
                $('#section_multi_search').empty();    
                $.each(data, function(key, item) {
                    $('#section_multi_search').append(`<option class='section_checkbox' id="section_checkbox_${item.Id}" value='${item.Id}' >${item.SectionName}</option>`)
                });  
                $('#section_multi_search').multiselect('rebuild');               
            });
        $.getJSON('/api/Departments/')
            .done(function(data) {
                $('#dept_multi_search').empty();
                $.each(data, function(key, item) {                    
                    $('#dept_multi_search').append(`<option class='department_checkbox' id="department_checkbox_${item.Id}" value='${item.Id}'>${item.DepartmentName}</option>`)
                });
                $('#dept_multi_search').multiselect('rebuild');
            });
        $.getJSON('/api/InCharges/')
            .done(function(data) {
                $('#incharge_multi_search').empty();
                $.each(data, function(key, item) {
                    $('#incharge_multi_search').append(`<option class='incharge_checkbox' id="incharge_checkbox_${item.Id}" value='${item.Id}'>${item.InChargeName}</option>`)
                });
                $('#incharge_multi_search').multiselect('rebuild');
            });
        $.getJSON('/api/Roles/')
            .done(function(data) {
                $('#role_multi_search').empty();
                $.each(data, function(key, item) {                    
                    $('#role_multi_search').append(`<option class='role_checkbox' id="role_checkbox_${item.Id}" value='${item.Id}'>${item.RoleName}</option>`)
                });
                $('#role_multi_search').multiselect('rebuild');
            });
        $.getJSON('/api/Explanations/')
            .done(function(data) {
                $('#explanation_multi_search').empty();
                $.each(data, function(key, item) {                    
                    $('#explanation_multi_search').append(`<option class='explanation_checkbox' id="explanation_checkbox_${item.Id}" value='${item.Id}'>${item.ExplanationName}</option>`)
                });
                $('#explanation_multi_search').multiselect('rebuild');
            });
        $.getJSON('/api/Companies/')
            .done(function(data) {
                $('#company_multi_search').empty();
                $.each(data, function(key, item) {
                    $('#company_multi_search').append(`<option class='comopany_checkbox' id="comopany_checkbox_${item.Id}" value='${item.Id}'>${item.CompanyName}</option>`)
                });
                $('#company_multi_search').multiselect('rebuild');
            });
    }
    

});
let previousAssignmentRow = 0;
let unitPrices = [];

function loadSingleAssignmentData(id) {
    //$("#modal_edit_name").modal('show');
    $.ajax({
        url: '/api/Utilities/AssignmentById/' + id,
        type: 'GET',
        dataType: 'json',
        success: function(assignmentData) {
            $('#edit_model_span').html(assignmentData.EmployeeName);
            $('#name_edit').val(assignmentData.EmployeeName);
            $('#memo_edit').val(assignmentData.Remarks);
            $('#row_id_hidden_edit').val(assignmentData.Id);
            $.getJSON('/api/sections/')
                .done(function(data) {
                    $('#section_edit').empty();
                    $('#section_edit').append(`<option value=''>Select Section</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.SectionId) {
                            $('#section_edit').append(`<option value='${item.Id}' selected>${item.SectionName}</option>`);
                        } else {
                            $('#section_edit').append(`<option value='${item.Id}'>${item.SectionName}</option>`);
                        }

                    });
                });

            $.getJSON(`/api/departments`)
                .done(function(data) {
                    $('#department_edit').empty();
                    $('#department_edit').append(`<option value=''>Select Department</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.DepartmentId) {
                            $('#department_edit').append(`<option value='${item.Id}' selected>${item.DepartmentName}</option>`);
                        } else {
                            $('#department_edit').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
                        }

                    });
                });

            $.getJSON('/api/incharges/')
                .done(function(data) {
                    $('#incharge_edit').empty();
                    $('#incharge_edit').append(`<option value=''>Select In-Charge</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.InchargeId) {
                            $('#incharge_edit').append(`<option value='${item.Id}' selected>${item.InChargeName}</option>`);
                        } else {
                            $('#incharge_edit').append(`<option value='${item.Id}'>${item.InChargeName}</option>`);
                        }

                    });
                });


            $.getJSON('/api/Roles/')
                .done(function(data) {
                    $('#role_edit').empty();
                    $('#role_edit').append(`<option value=''>Select Role</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.RoleId) {
                            $('#role_edit').append(`<option value='${item.Id}' selected>${item.RoleName}</option>`);
                        } else {
                            $('#role_edit').append(`<option value='${item.Id}'>${item.RoleName}</option>`);
                        }

                    });
                });


            $.getJSON('/api/Explanations/')
                .done(function(data) {
                    $('#explanation_edit').empty();
                    $('#explanation_edit').append(`<option value=''>Select Explanation</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.ExplanationId) {
                            $('#explanation_edit').append(`<option value='${item.Id}' selected>${item.ExplanationName}</option>`);
                        } else {
                            $('#explanation_edit').append(`<option value='${item.Id}'>${item.ExplanationName}</option>`);
                        }

                    });
                });

            $.getJSON('/api/Companies/')
                .done(function(data) {
                    $('#company_edit').empty();
                    $('#company_edit').append(`<option value=''>Select Company</option>`);
                    $.each(data, function(key, item) {
                        if (item.Id == assignmentData.CompanyId) {
                            $('#company_edit').append(`<option value='${item.Id}' selected>${item.CompanyName}</option>`);
                        } else {
                            $('#company_edit').append(`<option value='${item.Id}'>${item.CompanyName}</option>`);
                        }

                    });
                });
            $('#unitprice_edit').val(assignmentData.UnitPrice);
            if (assignmentData.CompanyName.toLowerCase() == "mw") {
                $('#grade_edit').val(assignmentData.GradePoint);
            } else {
                $('#grade_edit').val("");
            }
            $('#grade_edit_hidden').val(assignmentData.GradeId);

        },
        error: function() {
            //$('#add_name_table_1 tbody').empty();
        }
    });
}
function loadAssignmentRowData(id) {
    //$("#namelist_delete").modal('show');
    $('#namelist_inactive_rowid').val(id);
}
//save edit information
$('#add_name_edit').on('click', function() {
    var sectionId = $('#section_edit').find(":selected").val();
    var inchargeId = $('#incharge_edit').find(":selected").val();
    var departmentId = $('#department_edit').find(":selected").val();
    var roleId = $('#role_edit').find(":selected").val();
    var companyId = $('#company_edit').find(":selected").val();
    var explanationId = $('#explanation_edit').find(":selected").val();
    var unitPrice = $('#unitprice_edit').val();
    var gradeId = $('#grade_edit_hidden').val();
    var rowId = $('#row_id_hidden_edit').val();
    var remarks = $('#memo_edit').val();

    var dataSingleAssignmentUpdate = {
        Id: rowId,
        SectionId: sectionId,
        DepartmentId: departmentId,
        InchargeId: inchargeId,
        RoleId: roleId,
        ExplanationId: explanationId,
        CompanyId: companyId,
        UnitPrice: unitPrice,
        GradeId: gradeId,
        Remarks: remarks,
    };

    $.ajax({
        url: '/api/Employees',
        type: 'PUT',
        async: false,
        dataType: 'json',
        data: dataSingleAssignmentUpdate,
        success: function(data) {
            //ToastMessageSuccess(data);
            alert("success");
            location.reload();
        },
        error: function(data) {
            alert(data.responseJSON.Message);
        }
    });



});
//edit employee assignment: compare grade with company name
$('#company_edit').on('change', function() {
    var _companyName = $("#company_edit option:selected").text();
    var _unitPrice = $("#unitprice_edit").val();

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            $('#grade_edit_hidden').val(data.Id);
            if (_companyName.toLowerCase() == "mw") {
                $('#grade_edit').val(data.SalaryGrade);
            } else {
                $('#grade_edit').val('');
            }
        },
        error: function() {
            $('#grade_edit').val('');
            $('#grade_edit_hidden').val('');
        }
    });
});
//edit employee assignment: compare grade with unit price
$('#unitprice_edit').on('change', function() {
    var _unitPrice = $(this).val();
    var _companyName = $("#company_edit option:selected").text();

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            $('#grade_edit_hidden').val(data.Id);
            if (_companyName.toLowerCase() == "mw") {
                $('#grade_edit').val(data.SalaryGrade);
            } else {
                $('#grade_edit').val('');
            }
        },
        error: function() {
            $('#grade_edit').val('');
            $('#grade_edit_hidden').val('');
        }
    });
});
//Namelist Inactive Assignment
$('#namelist_inactive').on('click', function() {
    var rowId = $('#namelist_inactive_rowid').val();

    $.ajax({
        url: `/api/Employees/RemoveAssignment/${rowId}`,
        type: 'PUT',
        dataType: 'json',
        success: function(data) {
            ToastMessageSuccess(data);
            location.reload();
        },
        error: function() {
            alert("Error please try again");
        }
    });
});
function loadSingleAssignmentDataForExistingEmployee(employeeName) {
    $('#add_name_table_2 thead .sub_thead').remove();
    $('#add_name_table_2 tbody').empty();
    $('#add_name_add').css('display', 'inline-block');
    console.log(employeeName);
    $.ajax({
        url: `/api/Utilities/GetEmployeesByName/${employeeName}`,
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(assignmentData) {
            console.log(assignmentData);
            $('#fixed_hidden_name').val(assignmentData[0].EmployeeName);

            var count = 1;
            $.each(assignmentData, function(index, item) {
                // console.log(item);
                unitPrices.push(item.UnitPrice);
                previousAssignmentRow = item.SubCode;
                tr += '<tr class="sub_thead">';
                tr += `<td>${item.SubCode}</td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 141px;' value='${item.EmployeeName}' readonly/></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 47px;' value='${item.SubCode}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.Remarks}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 87px;' value='${item.SectionName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.DepartmentName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.InchargeName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 89px;' value='${item.RoleName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width: 190px;' value='${item.ExplanationName}' readonly /></td>`;
                tr += `<td><input type='text' class=" col-12" style='width:81px;'   value='${item.CompanyName}' readonly /></td>`;
                if (item.CompanyName.toLowerCase() == "mw") {
                    tr += `<td><input type='text' class=" col-12" value='${item.GradePoint}' readonly style='width: 55px;' /></td>`;
                } else {
                    tr += `<td><input type='text' class=" col-12" value='' readonly style='width: 55px;' /></td>`;
                }
                tr += `<td><input type='text' class="col-12" id="add_name_unit_price_${count}" style='width: 72px;' value='${item.UnitPrice}' readonly /></td>`;
                if (count == 1) {
                    $("#unit_price_first_project_hid").val(item.UnitPrice);
                    tr += `<td rowspan='${assignmentData.length}' style='text-align:center;'><a href='javascript:void(0);' onClick="addNew()"> <i class="fa fa-plus" aria-hidden="true"></i></a></td>`;
                }

                tr += '</tr>';

                //console.log(tr);
                count++;
                $('#add_name_table_2 thead').append(tr);
                var tr = "";
            });

        },
        error: function() {

        }
    });
}
function addNew() {
    //
    previousAssignmentRow++;
    var text = `
    <tr data-id="${previousAssignmentRow}" id="row_${previousAssignmentRow}">
        <td>${previousAssignmentRow}</td>
        <td><input class=" col-12" id="identity_row_${previousAssignmentRow}" style='width: 141px;' value="${$('#fixed_hidden_name').val()}" readonly /></td>
        <td><input class="" value="${previousAssignmentRow}"  style='width: 47px;' id='sub_code_row_${previousAssignmentRow}' readonly /></td>
        <td><input class=" col-12" style='width: 89px;' id="memo_row_${previousAssignmentRow}" /></td>`;

    // for add new
    $.ajax({
        url: '/api/Sections/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            console.log(data);
            text += '<td>';
            text += `
                        <select id="section_row_${previousAssignmentRow}" style='width: 87px;'class=" col-12 section_row" onChange="changeVal(${previousAssignmentRow},this)">
                            <option value=''>Select Section</option>
                    `;
            $.each(data, function(key, item) {
                text += `<option value = '${item.Id}'> ${item.SectionName}</option>`;
            });
            text += '</select></td>';
        },
        error: function() {

        }
    });


    text += `
        <td>
            <select id="department_row_${previousAssignmentRow}" class=" col-12 department_row"></select>
            </td>
        `;

    $.ajax({
        url: '/api/incharges/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            console.log(data);
            text += '<td>';
            text += `
                        <select id="incharge_row_${previousAssignmentRow}" class=" col-12">
                            <option value=''>Select In-Charge</option>
                    `;
            $.each(data, function(key, item) {
                text += `<option value = '${item.Id}'> ${item.InChargeName}</option>`;
            });
            text += '</select></td>';
        },
        error: function() {

        }
    });

    $.ajax({
        url: '/api/roles/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            console.log(data);
            text += '<td>';
            text += `
                        <select id="role_row_${previousAssignmentRow}" class=" col-12">
                            <option value=''>Select Role</option>
                    `;
            $.each(data, function(key, item) {
                text += `<option value = '${item.Id}'> ${item.RoleName}</option>`;
            });
            text += '</select></td>';
        },
        error: function() {

        }
    });

    $.ajax({
        url: '/api/Explanations/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            console.log(data);
            text += '<td>';
            text += `
                        <select id="explain_row_${previousAssignmentRow}" style='width: 190px;' class=" col-12">
                            <option value=''>Select Role</option>
                    `;
            $.each(data, function(key, item) {
                text += `<option value = '${item.Id}'> ${item.ExplanationName}</option>`;
            });
            text += '</select></td>';
        },
        error: function() {

        }
    });

    $.ajax({
        url: '/api/Companies/',
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            console.log(data);
            text += '<td>';
            text += `
                        <select id="company_row_${previousAssignmentRow}" style='width:81px;' class=" col-12 add_company" onchange="LoadGradeValue(this);">
                            <option value=''>Select Role</option>
                    `;
            $.each(data, function(key, item) {
                text += `<option value = '${item.Id}'> ${item.CompanyName}</option>`;
            });
            text += '</select></td>';
        },
        error: function() {

        }
    });
    text += `
            <td><input class=" col-12" id="grade_row_${previousAssignmentRow}" readonly style='width: 55px;'/><input type='hidden' id='grade_row_new_${previousAssignmentRow}' /></td>
            <td><input class=" col-12" id="unitprice_row_${previousAssignmentRow}" style='width: 72px;' onChange="unitPriceChange(${previousAssignmentRow},this)"/></td>
            <td style='text-align:center;'><a href='javascript:void(0);' id='remove_row_${previousAssignmentRow}' onClick="removeTr(${previousAssignmentRow})"><i class="fa fa-trash-o" aria-hidden="true"></i></a></td>
    </tr>

    `;

    $(`#remove_row_${previousAssignmentRow - 1}`).css('display', 'none');

    $('#add_name_table_2 tbody').append(text);

}

function removeTr(rowId) {
    var rowElement = $('#row_' + rowId);
    rowElement.remove();
    previousAssignmentRow--;
    $(`#remove_row_${previousAssignmentRow}`).css('display', 'block');

}


function onSave() {
    const rows = document.querySelectorAll("#add_name_table_2 > tbody > tr");
    //$('#add_name_add').css('display','none');
    console.log(rows);
    $.each(rows, function(index, data) {
        var rowId = data.dataset.id;

        var employeeName = $('#identity_row_' + rowId).val();
        var sectionId = $('#section_row_' + rowId).find(":selected").val();
        var departmentId = $('#department_row_' + rowId).find(":selected").val();
        var inchargeId = $('#incharge_row_' + rowId).find(":selected").val();
        var roleId = $('#role_row_' + rowId).find(":selected").val();
        var companyId = $('#company_row_' + rowId).find(":selected").val();
        var explanationId = $('#explain_row_' + rowId).find(":selected").val();
        var unitPrice = $('#unitprice_row_' + rowId).val();
        var gradeId = $('#grade_row_new_' + rowId).val();
        var remarks = $('#memo_row_' + rowId).val();
        var subCode = $('#sub_code_row_' + rowId).val();



        var data = {
            EmployeeName: employeeName,
            SectionId: sectionId,
            DepartmentId: departmentId,
            InchargeId: inchargeId,
            RoleId: roleId,
            ExplanationId: explanationId,
            CompanyId: companyId,
            UnitPrice: unitPrice,
            GradeId: gradeId,
            Remarks: remarks,
            SubCode: subCode,

        };

        console.log(data);

        $.ajax({
            url: '/api/Employees',
            type: 'POST',
            async: false,
            dataType: 'json',
            data: data,
            success: function(data) {

                Command: toastr["success"](data, "Success")

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
                alert('Success');
                location.reload();
                //$('#modal_add_name_new').modal('toggle');
            },
            error: function() {
                alert("Error please try again");
                //$('#modal_add_name_new').modal('toggle');
            }
        });
    });

}
function changeVal(rowId, element) {

    var sectionId = $(element).val();

    $.ajax({
        url: `/api/utilities/DepartmentsBySection/${sectionId}`,
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            $('#department_row_' + rowId).empty();
            $('#department_row_' + rowId).append(`<option value=''>Select Department</option>`);
            $.each(data, function(key, item) {
                $('#department_row_' + rowId).append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        },
        error: function() {

        }
    });


}
function unitPriceChange(rowId, element) {
    let _unitPrice = $(element).val();
    $("#add_name_unit_price_hidden").val(_unitPrice);
    $("#add_name_row_no_hidden").val(rowId);
    let _unitPrideFirstProject = $("#unit_price_first_project_hid").val();
    let _unitPrideFirstProjectWithComma = $("#add_name_unit_price_1").val();

    var _rowId = $("#add_name_row_no_hidden").val();
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    console.log("_companyName: " + _companyName);
    //var _unitPrice = $("#add_name_unit_price_hidden").val();

    console.log("_unitPrice: " + _unitPrice);


    //if (!unitPrices.includes(_unitPrice)) {
    if (_unitPrideFirstProject != _unitPrice && _unitPrideFirstProjectWithComma != _unitPrice) {
        $('#identity_row_' + rowId).css('color', 'red');
        $('#sub_code_row_' + rowId).css('color', 'red');
    } else {
        $('#identity_row_' + rowId).css('color', '');
        $('#sub_code_row_' + rowId).css('color', '');
    }
    //else {
    //    $('#identity_row_' + rowId).css('color', 'red');
    //    $('#sub_code_row_' + rowId).css('color', 'red');
    //}

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            console.log(data);
            $('#grade_edit_hidden').val(data.Id);
            console.log("_companyName.toLowerCase(): " + _companyName.toLowerCase());
            console.log("_companyName.toLowerCase(): " + typeof(_companyName));
            console.log("_rowId: " + _rowId);
            //if (_companyName.toLowerCase == 'mw') {
            if (_companyName.toLowerCase().indexOf("mw") > 0) {
                $('#grade_row_' + rowId).attr('data-id', data.Id);
                $('#grade_row_new_' + rowId).val(data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
            } else {
                $('#grade_row_' + rowId).attr('data-id', data.Id);
                $('#grade_row_new_' + rowId).val(data.Id);
                $('#grade_row_' + _rowId).val('');
            }


            //$('#grade_row_' + rowId).val(data.SalaryGrade);
        },
        error: function() {
            $('#grade_row_' + rowId).attr('data-id', '');
            $('#grade_row_' + rowId).val('');
            $('#grade_row_new_' + rowId).val('');
            alert("Invalid Grade Id!!!");
        }
    });
}

function changeVal(rowId, element) {

    var sectionId = $(element).val();

    $.ajax({
        url: `/api/utilities/DepartmentsBySection/${sectionId}`,
        type: 'GET',
        async: false,
        dataType: 'json',
        success: function(data) {
            $('#department_row_' + rowId).empty();
            $('#department_row_' + rowId).append(`<option value=''>Select Department</option>`);
            $.each(data, function(key, item) {
                $('#department_row_' + rowId).append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
            });
        },
        error: function() {

        }
    });


}

function LoadGradeValue(sel) {
    var _rowId = $("#add_name_row_no_hidden").val();
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    console.log("_companyName: " + _companyName);
    var _unitPrice = $("#add_name_unit_price_hidden").val();

    console.log("_unitPrice: " + _unitPrice);

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        //data: {
        //    unitPrice: _unitPrice
        //},
        success: function(data) {
            $('#grade_edit_hidden').val(data.Id);
            console.log("_companyName.toLowerCase(): " + _companyName.toLowerCase());
            console.log("_companyName.toLowerCase(): " + typeof(_companyName));
            console.log("_rowId: " + _rowId);
            //if (_companyName.toLowerCase == 'mw') {
            if (_companyName.toLowerCase().indexOf("mw") > 0) {
                //$('#grade_new_hidden').val(data.Id);
                $('#grade_row_' + _rowId).attr('data-id', data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
                $('#grade_row_new_' + _rowId).val(data.Id);
            } else {
                $('#grade_row_' + _rowId).attr('data-id', '');
                $('#grade_row_' + _rowId).val('');
                $('#grade_row_new_' + _rowId).val('');
            }
        },
        error: function() {
            $('#grade_row_' + _rowId).attr('data-id', '');
            $('#grade_row_' + _rowId).val('');
            $('#grade_row_new_' + _rowId).val('');
        }
    });
}
function GetMultiSearch_AjaxData() {  
    var employeeName = $('#name_search').val();
    var sectionCheck = $('#section_multi_search').val();  
    var departmentCheck = $('#dept_multi_search').val();  
    var inchargeCheck = $('#incharge_multi_search').val();  
    var roleCheck = $('#role_multi_search').val();  
    var explanationCheck = $('#explanation_multi_search').val();  
    var companyCheck = $('#company_multi_search').val();  

    var data_info = {
        EmployeeName: employeeName,
        Sections: sectionCheck,
        Departments: departmentCheck,
        Incharges: inchargeCheck,
        Roles: roleCheck,
        Explanations: explanationCheck,
        Companies: companyCheck

    };

    $.ajax({
        url: `/api/utilities/SearchMultipleEmployee`,
        type: 'POST',
        dataType: 'json',
        data: data_info,
        success: function(data) {
            $('#employee_list_search_results').empty();
            NameList_DatatableLoad(data);
            //NameListSort_OnChange();
        },
        error: function() {}
    });
    //return data_info;
}
function NameList_DatatableLoad(data) {
    var tempCompanyName = "";
    var tempSubCode = "";
    var redMark = "";

    $('#namelist').DataTable({
        destroy: true,
        data: data,
        ordering: true,
        orderCellsTop: true,
        pageLength: 10,
        searching: false,
        bLengthChange: false,
        //pagingType: 'full_numbers',
        //pagingType: 'input',
        //dom: 'lifrtip',
        
        columns: [
            //{
            //    data: 'MarkedAsRed',
            //    render: function (markedAsRed) {
            //        redMark = markedAsRed;
            //        return null;
            //    }
            //},
            {
                data: 'EmployeeNameWithCodeRemarks',
                render: function(employeeNameWithCodeRemarks) {
                    var splittedString = employeeNameWithCodeRemarks.split('$');
                    if (splittedString[2] == 'true') {
                        return `<span style='color:red;' class='namelist_addname' onClick="loadSingleAssignmentDataForExistingEmployee('${splittedString[0]}')" data-toggle="modal" data-target="#modal_add_name">${splittedString[1]}</span>`;
                    } else {
                        return `<span class='namelist_addname' onClick="loadSingleAssignmentDataForExistingEmployee('${splittedString[0]}')" data-toggle="modal" data-target="#modal_add_name">${splittedString[1]}</span>`;
                    }

                }
            },
            {
                data: 'SectionName'
            },
            {
                data: 'DepartmentName'
            },
            {
                data: 'InchargeName'
            },
            {
                data: 'RoleName'
            },
            {
                data: 'ExplanationName'
            },
            {
                data: 'CompanyName',
                render: function(companyName) {
                    tempCompanyName = companyName;
                    return companyName;
                }
            },
            {
                data: 'GradePoint',
                render: function(grade) {
                    if (tempCompanyName.toLowerCase() == "mw") {
                        return grade;
                    } else {
                        return "<span style='display:none;'>" + grade + "</span>";
                    }
                }
            },
            {
                data: 'UnitPrice'
            },
            {
                data: 'Id',
                render: function(Id) {
                    if (tempCompanyName.toLowerCase() == 'mw') {
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle ="modal"  data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    } else {
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle ="modal"  data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    }
                },
                orderable: false
            }
        ]
    });
}

//search by employee name
$(document).on('change', '#name_search', function() {
    GetMultiSearch_AjaxData();
});

$('#section_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});

//search by multi department
$('#dept_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});

//search by multi incharge
$('#incharge_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});

//search by multi roles
$('#role_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});

//search by multi explanations
$('#explanation_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});

//search by multi companies
$('#company_multi_search').change(function() {
    GetMultiSearch_AjaxData();
});