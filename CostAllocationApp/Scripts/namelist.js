$(document).ready(function() {
    var count=1;
    $('#namelist thead tr:eq(1) th').each( function () {
        if(count == 1){
            var title = $(this).text();
            $(this).html( '<input type="text" placeholder="Search '+title+'" class="column_search" />' );
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
            orderCellsTop: true,
            columnDefs: [
                //{ width: 10, targets: 0 }
            ],
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
                            return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-bs-toggle="modal"  data-bs-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                        } else {
                            return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-bs-toggle="modal"  data-bs-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                        }
                    },
                    orderable: false
                }
            ]
        });
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
function loadSingleAssignmentData(id) {
    $("#modal_edit_name").modal('show');
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
    $('#namelist_inactive_rowid').val(id);
}