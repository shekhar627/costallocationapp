$(document).ready(function() {
    GetNameList();

    //search by employee name
    $(document).on('change', '#name_search', function() {
        $("#multi_search_clicked_id").val('');
        GetMultiSearch_AjaxData('name');
    });

    //searcy by multi sections
    $(document).on('click', '#sectionChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('section');
    });

    //search by multi department
    $(document).on('click', '#departmentChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('dept');
    });

    //search by multi incharge
    $(document).on('click', '#inchargeChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('incharge');
    });

    //search by multi roles
    $(document).on('click', '#RoleChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('role');
    });

    //search by multi explanations
    $(document).on('click', '#ExplanationChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('explanation');
    });

    //search by multi companies
    $(document).on('click', '#CompanyChks input[type="checkbox"]', function() {
        $("#multi_search_clicked_id").val('');
        var clicked_checkbox = $(this).attr("id");
        $("#multi_search_clicked_id").val(clicked_checkbox);
        GetMultiSearch_AjaxData('company');
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
                ToastMessageSuccess(data);
                location.reload();
            },
            error: function(data) {
                alert(data.responseJSON.Message);
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

    //edit employee assignment: department dropdown fill aganist section
    $(document).on('change', '#section_edit', function() {
        var sectionId = $(this).val();
        $.getJSON(`/api/utilities/DepartmentsBySection/${sectionId}`)
            .done(function(data) {
                $('#department_edit').empty();
                $('#department_edit').append(`<option value=''>Select Department</option>`);
                $.each(data, function(key, item) {
                    $('#department_edit').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
                });
            });
    });

    //Section Registration Modal
    $('#section_registration').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    //Department Registration Modal
    $('#department_registration').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    //Incharge Registration Modal
    $('#in-charge_registration').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    //Role Registration Modal
    $('#role_registration_modal').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    //Explanation Registration Modal
    $('#explanation_modal').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    //Company Registration Modal
    $('#company_reg_modal').on('hidden.bs.modal', function() {
        var isPageLoad = $("#page_load_after_modal_close").val();
        if (isPageLoad == "yes") {
            GetListDropdownValue();
        }
    })

    var expanded = false;
});
//checked here
function GetNameList() {
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
            $('#sectionChks').empty();
            $('#sectionChks').append(`<label for='chk_sec_all'><input id="chk_sec_all" type="checkbox" checked data-checkwhat="chkSelect" value='sec_all'/>All</label>`)
            $.each(data, function(key, item) {
                $('#sectionChks').append(`<label for='section_checkbox_${item.Id}'><input class='section_checkbox' id="section_checkbox_${item.Id}"  type="checkbox" checked value='${item.Id}'/>${item.SectionName}</label>`)
            });
        });
    $.getJSON('/api/Departments/')
        .done(function(data) {
            $('#departmentChks').empty();
            $('#departmentChks').append(`<label for='chk_dept_all'><input id="chk_dept_all" type="checkbox" checked value='dept_all'/>All</label>`)
            $.each(data, function(key, item) {
                $('#departmentChks').append(`<label for='department_checkbox_${item.Id}'><input class='department_checkbox' id="department_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.DepartmentName}</label>`)
            });
        });
    $.getJSON('/api/InCharges/')
        .done(function(data) {
            $('#inchargeChks').empty();
            $('#inchargeChks').append(`<label for='chk_incharge_all'><input id="chk_incharge_all" type="checkbox" checked value='incharge_all'/>All</label>`);
            $.each(data, function(key, item) {
                $('#inchargeChks').append(`<label for='incharge_checkbox_${item.Id}'><input class='incharge_checkbox' id="incharge_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.InChargeName}</label>`)
            });
        });
    $.getJSON('/api/Roles/')
        .done(function(data) {
            $('#RoleChks').empty();
            $('#RoleChks').append(`<label for='chk_role_all'><input id="chk_role_all" type="checkbox" checked value='role_all'/>All</label>`);
            $.each(data, function(key, item) {
                $('#RoleChks').append(`<label for='role_checkbox_${item.Id}'><input class='role_checkbox' id="role_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.RoleName}</label>`)
            });
        });
    $.getJSON('/api/Explanations/')
        .done(function(data) {
            $('#ExplanationChks').empty();
            $('#ExplanationChks').append(`<label for='chk_explanation_all'><input id="chk_explanation_all" type="checkbox" checked value='explanation_all'/>All</label><br>`);
            $.each(data, function(key, item) {
                $('#ExplanationChks').append(`<label for='explanation_checkbox_${item.Id}'><input class='explanation_checkbox' id="explanation_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.ExplanationName}</label><br>`);

            });
        });
    $.getJSON('/api/Companies/')
        .done(function(data) {
            $('#CompanyChks').empty();
            $('#CompanyChks').append(`<label for='chk_comopany_all'><input id="chk_comopany_all" type="checkbox" checked value='comopany_all'/>All</label>`);
            $.each(data, function(key, item) {
                $('#CompanyChks').append(`<label for='comopany_checkbox_${item.Id}'><input class='comopany_checkbox' id="comopany_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.CompanyName}</label>`)
            });
        });
}

function NameList_DatatableLoad(data) {
    var tempCompanyName = "";
    var tempSubCode = "";
    var redMark = "";

    $('#name-list').DataTable({
        destroy: true,
        data: data,
        ordering: true,
        orderCellsTop: true,
        pageLength: 100,
        searching: false,
        bLengthChange: false,
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
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle="modal" data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    } else {
                        return `<td class='namelist_td Action'><a href="javascript:void(0);" id='edit_button' onClick="loadSingleAssignmentData(${Id})" data-toggle="modal" data-target="#modal_edit_name">Edit</a><a id='delete_button' href='javascript:void();' data-toggle='modal' data-target='#namelist_delete' onClick="loadAssignmentRowData(${Id})" assignment_id='${Id}'>Inactive</a></td>`;
                    }
                }
            }
        ]
    });
}

function GetMultiSearch_AjaxData(searchType) {
    let isSectionAllChk = $("#chk_sec_all").is(':checked');
    var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
    var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
    var isRoleAllCheck = $("#chk_role_all").is(':checked');
    var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
    var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');


    var all_check_uncheck = $("#multi_search_clicked_id").val();
    if (searchType.toLowerCase() == 'section') {
        if (all_check_uncheck.toLowerCase() == "chk_sec_all") {
            $(".section_checkbox").prop("checked", isSectionAllChk);
        } else {
            $("#chk_sec_all").prop("checked", false);
        }
    } else if (searchType.toLowerCase() == 'dept') {
        if (all_check_uncheck.toLowerCase() == "chk_dept_all") {
            $(".department_checkbox").prop("checked", isDepartmentAllCheck);
        } else {
            $("#chk_dept_all").prop("checked", false);
        }
    } else if (searchType.toLowerCase() == 'incharge') {
        if (all_check_uncheck.toLowerCase() == "chk_incharge_all") {
            $(".incharge_checkbox").prop("checked", isInChargeAllCheck);
        } else {
            $("#chk_incharge_all").prop("checked", false);
        }
    } else if (searchType.toLowerCase() == 'role') {
        if (all_check_uncheck.toLowerCase() == "chk_role_all") {
            $(".role_checkbox").prop("checked", isRoleAllCheck);
        } else {
            $("#chk_role_all").prop("checked", false);
        }
    } else if (searchType.toLowerCase() == 'explanation') {
        if (all_check_uncheck.toLowerCase() == "chk_explanation_all") {
            $(".explanation_checkbox").prop("checked", isExplanationAllCheck);
        } else {
            $("#chk_explanation_all").prop("checked", false);
        }
    } else if (searchType.toLowerCase() == 'company') {
        if (all_check_uncheck.toLowerCase() == "chk_comopany_all") {
            $(".comopany_checkbox").prop("checked", isCompanytAllCheck);
        } else {
            $("#chk_comopany_all").prop("checked", false);
        }
    }
    all_check_uncheck = "";

    var employeeName = "";
    var sectionCheck = [];
    var departmentCheck = [];
    var inchargeCheck = [];
    var roleCheck = [];
    var explanationCheck = [];
    var companyCheck = [];



    employeeName = $('#name_search').val();
    var sectionCheckedBoxes = $('#sectionChks input[type="checkbox"]:checked');
    var departmentCheckedBoxes = $('#departmentChks input[type="checkbox"]:checked');
    var inchargeCheckedBoxes = $('#inchargeChks input[type="checkbox"]:checked');
    var roleCheckedBoxes = $('#RoleChks input[type="checkbox"]:checked');
    var explanationCheckedBoxes = $('#ExplanationChks input[type="checkbox"]:checked');
    var companyCheckedBoxes = $('#CompanyChks input[type="checkbox"]:checked');

    if (!isSectionAllChk) {
        $.each(sectionCheckedBoxes, function(index, item) {
            sectionCheck.push(item.value);
        });
    }
    if (!isDepartmentAllCheck) {
        $.each(departmentCheckedBoxes, function(index, item) {
            departmentCheck.push(item.value);
        });
    }
    if (!isInChargeAllCheck) {
        $.each(inchargeCheckedBoxes, function(index, item) {
            inchargeCheck.push(item.value);
        });
    }
    if (!isRoleAllCheck) {
        $.each(roleCheckedBoxes, function(index, item) {
            roleCheck.push(item.value);
        });
    }
    if (!isExplanationAllCheck) {
        $.each(explanationCheckedBoxes, function(index, item) {
            explanationCheck.push(item.value);
        });
    }
    if (!isCompanytAllCheck) {
        $.each(companyCheckedBoxes, function(index, item) {
            companyCheck.push(item.value);
        });
    }

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
            NameListSort_OnChange();
        },
        error: function() {}
    });
    //return data_info;
}

function NameListSort(sort_asc, sort_desc) {
    var nameAsc = $('#' + sort_asc).css('display');
    if (nameAsc == 'inline-block') {
        $('#' + sort_asc).css('display', 'none');
        $('#' + sort_desc).css('display', 'inline-block');
    } else {
        $('#' + sort_asc).css('display', 'inline-block');
        $('#' + sort_desc).css('display', 'none');
    }
}

function SectionCheck() {
    DismissOtherDropdown("section");

    var checkboxes = document.getElementById("sectionChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        //$(this).find("#sectionChks").slideToggle("fast");
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

function DismissOtherDropdown(requestType) {
    var section_display = "";
    if (requestType.toLowerCase() != "section") {
        section_display = $('#sectionChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("sectionChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "dept") {
        section_display = "";
        section_display = $('#departmentChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("departmentChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "incharge") {
        section_display = "";
        section_display = $('#inchargeChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("inchargeChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "role") {
        section_display = "";
        section_display = $('#RoleChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("RoleChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "explanation") {
        section_display = "";
        section_display = $('#ExplanationChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("ExplanationChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }

    if (requestType.toLowerCase() != "company") {
        section_display = "";
        section_display = $('#CompanyChks');
        if (section_display.css("display") == "block") {
            let checkboxes = document.getElementById("CompanyChks");
            checkboxes.style.display = "none";
            expanded = false;
        }
    }
}

function NameListSort_OnChange() {
    $('#name_asc').css('display', 'inline-block');
    $('#name_desc').css('display', 'none');

    $('#section_asc').css('display', 'inline-block');
    $('#section_desc').css('display', 'none');

    $('#department_asc').css('display', 'inline-block');
    $('#department_desc').css('display', 'none');

    $('#incharge_asc').css('display', 'inline-block');
    $('#incharge_desc').css('display', 'none');

    $('#role_asc').css('display', 'inline-block');
    $('#role_desc').css('display', 'none');

    $('#explanation_asc').css('display', 'inline-block');
    $('#explanation_desc').css('display', 'none');

    $('#company_asc').css('display', 'inline-block');
    $('#company_desc').css('display', 'none');

    $('#grade_asc').css('display', 'inline-block');
    $('#grade_desc').css('display', 'none');

    $('#unit_asc').css('display', 'inline-block');
    $('#unit_desc').css('display', 'none');
}

function DepartmentCheck() {
    DismissOtherDropdown("dept");

    var checkboxes = document.getElementById("departmentChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

function InchargeCheck() {
    DismissOtherDropdown("incharge");
    var checkboxes = document.getElementById("inchargeChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

function RoleCheck() {
    DismissOtherDropdown("role");
    var checkboxes = document.getElementById("RoleChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

function ExplanationCheck() {
    DismissOtherDropdown("explanation");
    var checkboxes = document.getElementById("ExplanationChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

function CompanyCheck() {
    DismissOtherDropdown("company");
    var checkboxes = document.getElementById("CompanyChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

let previousAssignmentRow = 0;
let unitPrices = [];

function loadSingleAssignmentData(id) {

    $.ajax({
        url: '/api/Utilities/AssignmentById/' + id,
        type: 'GET',
        dataType: 'json',
        success: function(assignmentData) {

            console.log(assignmentData);
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

function loadAssignmentRowData(id) {
    $('#namelist_inactive_rowid').val(id);
}

var expanded = false;
$(document).on("click", function(event) {
    var $trigger = $(".multiselect");
    //alert("trigger: "+$trigger);
    if ($trigger !== event.target && !$trigger.has(event.target).length) {
        expanded = false;
        $(".commonselect").slideUp("fast");
    }
});