var globalSearchObject = '';
var globalPreviousValue='0.0';
var globalPreviousId = '';

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
//var expanded = false;
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
function DepartmentCheck() {
    DismissOtherDropdown("dept");
    // section_display = $('#sectionChks');
    // if (section_display.css("display") == "block") {
    //     let checkboxes = document.getElementById("sectionChks");
    //     checkboxes.style.display = "none";
    //     expanded = false;
    // }

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

function LoadGradeValue(sel) {
    var _rowId = $("#add_name_row_no_hidden").val();
    var _companyName = $('#company_row_' + _rowId).find(":selected").text();
    var _unitPrice = $("#add_name_unit_price_hidden").val();

    $.ajax({
        url: `/api/utilities/CompareGrade/${_unitPrice}`,
        type: 'GET',
        dataType: 'json',
        //data: {
        //    unitPrice: _unitPrice
        //},
        success: function (data) {
            $('#grade_edit_hidden').val(data.Id);
            //if (_companyName.toLowerCase == 'mw') {
            if (_companyName.toLowerCase().indexOf("mw") > 0) {
                //$('#grade_new_hidden').val(data.Id);
                $('#grade_row_' + _rowId).val(data.SalaryGrade);
            } else {
                $('#grade_row_' + _rowId).val('');
            }
        },
        error: function () {
            $('#grade_row_' + _rowId).val('');
        }
    });
}

function NameListSort(sort_asc, sort_desc) {
    var nameAsc = $('#' + sort_asc).css('display');
    if (nameAsc == 'inline-block') {
        $('#' + sort_asc).css('display', 'none');
        $('#' + sort_desc).css('display', 'inline-block');
    }
    else {
        $('#' + sort_asc).css('display', 'inline-block');
        $('#' + sort_desc).css('display', 'none');
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
                render: function (employeeNameWithCodeRemarks) {
                    var splittedString = employeeNameWithCodeRemarks.split('$');
                    if (splittedString[2] == 'true') {
                        return `<span style='color:red;' class='namelist_addname' onClick="loadSingleAssignmentDataForExistingEmployee('${splittedString[0]}')" data-toggle="modal" data-target="#modal_add_name">${splittedString[1]}</span>`;
                    }
                    else {
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
                render: function (companyName) {
                    tempCompanyName = companyName;
                    return companyName;
                }
            },
            {
                data: 'GradePoint',
                render: function (grade) {
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
                render: function (Id) {
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

//$("#section_registration").on('hide', function () {
//    window.location.reload();
//});

function RetainRegistrationValue() {
    let name = $("#hid_name").val();
    let memo = $("#hid_memo").val(memo);
    let sectionId = $("#hid_sectionId").val(sectionId);
    let departmentId = $("#hid_departmentId").val(departmentId);
    let inchargeId = $("#hid_inchargeId").val(inchargeId);
    let roleId = $("#hid_roleId").val(roleId);
    let explantionId = $("#hid_explanationId").val(explantionId);
    let companyid = $("#hid_companyId").val();
    let companyName = $("#hid_companyName").val();
    let grade = $("#hid_grade").val();
    //let unitPrice = $("#hid_gradeId").val(unitPrice);
    let unitPrice = $("#hid_unitPrice").val();

    $("#identity_new").val(name);
    $("#memo_new").val(memo);
    // $('#section_new').find(":selected").val();
    // $('#department_new').find(":selected").val();
    // $('#incharge_new').find(":selected").val();
    // $('#role_new').find(":selected").val();
    // $('#explanation_new').find(":selected").val();
    // $('#company_new').find(":selected").val();
    if (companyName.toLowerCase() == 'mw') {
        $("#grade_new").val(grade);
    } else {
        $("#grade_new").val('');
    }

    $("#unitprice_new").val(unitPrice);

}


$('#department_modal_href').click(function () {
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#section_list').empty();
            $('#section_list').append(`<option value=''>Select Section</option>`)
            $.each(data, function (key, item) {
                $('#section_list').append(`<option value='${item.Id}'>&nbsp;&nbsp; ${item.SectionName}</option>`)
            });
        });
});

function ForecastSearchDropdownInLoad() {
    $.getJSON('/api/sections/')
        .done(function (data) {
            $('#sectionChks').empty();
            $('#sectionChks').append(`<label for='chk_sec_all'><input id="chk_sec_all" type="checkbox" checked data-checkwhat="chkSelect" value='sec_all'/>All</label>`)
            $.each(data, function (key, item) {
                $('#sectionChks').append(`<label for='section_checkbox_${item.Id}'><input class='section_checkbox' id="section_checkbox_${item.Id}"  type="checkbox" checked value='${item.Id}'/>${item.SectionName}</label>`)
            });
        });
    $.getJSON('/api/Departments/')
        .done(function (data) {
            $('#departmentChks').empty();
            $('#departmentChks').append(`<label for='chk_dept_all'><input id="chk_dept_all" type="checkbox" checked value='dept_all'/>All</label>`)
            $.each(data, function (key, item) {
                $('#departmentChks').append(`<label for='department_checkbox_${item.Id}'><input class='department_checkbox' id="department_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.DepartmentName}</label>`)
            });
        });
    $.getJSON('/api/InCharges/')
        .done(function (data) {
            $('#inchargeChks').empty();
            $('#inchargeChks').append(`<label for='chk_incharge_all'><input id="chk_incharge_all" type="checkbox" checked value='incharge_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#inchargeChks').append(`<label for='incharge_checkbox_${item.Id}'><input class='incharge_checkbox' id="incharge_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.InChargeName}</label>`)
            });
        });
    $.getJSON('/api/Roles/')
        .done(function (data) {
            $('#RoleChks').empty();
            $('#RoleChks').append(`<label for='chk_role_all'><input id="chk_role_all" type="checkbox" checked value='role_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#RoleChks').append(`<label for='role_checkbox_${item.Id}'><input class='role_checkbox' id="role_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.RoleName}</label>`)
            });
        });
    $.getJSON('/api/Explanations/')
        .done(function (data) {
            $('#ExplanationChks').empty();
            $('#ExplanationChks').append(`<label for='chk_explanation_all'><input id="chk_explanation_all" type="checkbox" checked value='explanation_all'/>All</label><br>`);
            $.each(data, function (key, item) {
                $('#ExplanationChks').append(`<label for='explanation_checkbox_${item.Id}'><input class='explanation_checkbox' id="explanation_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.ExplanationName}</label><br>`);

            });
        });
    $.getJSON('/api/Companies/')
        .done(function (data) {
            $('#CompanyChks').empty();
            $('#CompanyChks').append(`<label for='chk_comopany_all'><input id="chk_comopany_all" type="checkbox" checked value='comopany_all'/>All</label>`);
            $.each(data, function (key, item) {
                $('#CompanyChks').append(`<label for='comopany_checkbox_${item.Id}'><input class='comopany_checkbox' id="comopany_checkbox_${item.Id}" type="checkbox" checked value='${item.Id}'/>${item.CompanyName}</label>`)
            });
        });
}
function CheckPreviousManMonthIdValuePoint(e){
    let clickedId = e.id;
    console.log("clickedId: "+clickedId);
    if(globalPreviousId == ''){
        globalPreviousId = clickedId;
    }

    if(globalPreviousId == clickedId){
        let pointValue = $("#" + e.id).val();
        if ((isNaN(pointValue) || pointValue == undefined || pointValue == '') && globalPreviousValue >0){
            globalPreviousValue = globalPreviousValue;
        }else{
            globalPreviousValue = $("#" + e.id).val();
        }    
    
        if (e.value <= 0) {
            //globalPreviousValue='';
            $("#" + e.id).val('');
        }
    }else{
                
        let pointValue = $("#" + globalPreviousId).val();
        if ((isNaN(pointValue) || pointValue == undefined || pointValue == '') && globalPreviousValue >0){
            globalPreviousValue = globalPreviousValue;
        }else{
            globalPreviousValue = $("#" + globalPreviousId).val();
        }    
    
        if (e.value <= 0) {
            //globalPreviousValue='';
            $("#" + e.id).val('');
        }
        globalPreviousId = clickedId;
    
    }
    console.log("clickedId: "+clickedId);
    console.log("globalPreviousId: "+globalPreviousId);
    console.log("globalPreviousValue: "+globalPreviousValue);
}

function checkPoint_click(e) {
    console.log("click: "+globalPreviousValue);
    let pointValue = $("#" + e.id).val();
    let comparePreviousGlobalValue = parseFloat(globalPreviousValue);
    if ((isNaN(pointValue) || pointValue == undefined || pointValue == '') && comparePreviousGlobalValue > 0) {
        globalPreviousValue = globalPreviousValue;
    }
    else if (comparePreviousGlobalValue > 0) {
        globalPreviousValue = globalPreviousValue;
    }
    else {
        //globalPreviousValue = $("#" + e.id).val();
        globalPreviousValue = '0.0';
    }    

    if (e.value <= 0) {
        //globalPreviousValue='';
        $("#" + e.id).val('');
        // $("#" + e.id).val('0.0');
    }
}

function checkPoint_onmouseover(e) {
    let pointValue = $("#" + e.id).val();
    if ((isNaN(pointValue) || pointValue == undefined || pointValue == '') && globalPreviousValue >0){
        globalPreviousValue = globalPreviousValue;
    }else{
        globalPreviousValue = $("#" + e.id).val();
    }
    
    console.log(globalPreviousValue);
}
function checkPoint(element) {

    var pointValue = $(element).val();
    if (isNaN(pointValue) || pointValue == undefined || pointValue == '') {        
        $(element).val(globalPreviousValue);
    }
    else {
        if ((pointValue > 1 || pointValue < 0) ) {
            alert('total month point can not be grater than 1 Or less than 0');
            $(element).val(globalPreviousValue);
        }
        else{
            $(element).val(pointValue);
        }
    }


    var totalMonthPoint = 0;
    var sameNameTr = [];
    var tr = element.closest('tr');
    //var trId = $(tr).data('rowid');
    var trName = $(tr).find('td').eq(0).children('span').data('name');
    var allTr = $('#forecast_table > tbody > tr');
    //var monthNumber = $(element).data('month');
    var columnIndex = $(element).parent().index();

    $.each(allTr, function (index, value) {
        var tempTrName = $(value).find('td').eq(0).children('span').data('name');
        if (tempTrName == trName) {
            var monthPoint = $(value).find('td').eq(columnIndex).children('input').val();
            if (monthPoint == '' || monthPoint == NaN) {
                totalMonthPoint += 0;
            }
            else {
                totalMonthPoint += parseFloat(monthPoint);
            }
            sameNameTr.push(value);
        }
    });
     if (totalMonthPoint > 1) {
         alert('total month point can not be grater than 1');
         $(element).val(globalPreviousValue);
     }
}
function LoaderShow(){
    $("#forecast_table_wrapper").css("display", "none");
    $("#loading").css("display", "block");
}
function LoaderHide(){
    $("#forecast_table_wrapper").css("display", "block");
    $("#loading").css("display", "none");
}

function LoadForecastData(){
    console.log("saved-3");
    if (globalSearchObject == '') {
        return;
    }
    else {
        $.ajax({
            url: `/api/utilities/SearchForecastEmployee`,
            contentType: 'application/json',
            type: 'GET',
            async: true,
            dataType: 'json',
            data: globalSearchObject,            
            success: function (data) {
                LoaderHide();
                $('#forecast_table>thead').empty();
                $('#forecast_table>tbody').empty();
                $('#forecast_table>thead').append(`
                    <tr>
                        <th id="forecast_name">Name </th>
                        <th id="forecast_section">Section </th>
                        <th id="forecast_department">Department </th>
                        <th id="forecast_incharge">In-Charge </th>
                        <th id="forecast_role">Role </th>
                        <th id="forecast_explanation">Explanation </th>
                        <th id="forecast_company">Company Name </i></th>
                        <th id="forecast_grade">Grade </th>
                        <th id="forecast_unitprice"><span>Unit Price</span> </th>
                        <th>10月</th>
                        <th>11月</th>
                        <th>12月</th>
                        <th>1月</th>
                        <th>2月</th>
                        <th>3月</th>
                        <th>4月</th>
                        <th>5月</th>
                        <th>6月</th>
                        <th>7月</th>
                        <th>8月</th>
                        <th>9月</th>
                        <th>Oct</th>
                        <th>Nov</th>
                        <th>Dec</th>
                        <th>Jan</th>
                        <th>Feb</th>
                        <th>Mar</th>
                        <th>Apr</th>
                        <th>May</th>
                        <th>Jun</th>
                        <th>Jul</th>
                        <th>Aug</th>
                        <th>Sep</th>
                    </tr>
                `);
                //Forecast_DatatableLoad(data);
                var _id = 0;
                var _companyName = '';
                var _oct = 0;
                var _octTotal = 0;

                $('#forecast_table').DataTable({
                    destroy: true,
                    data: data,
                    searching: false,
                    bLengthChange: false,
                    autoWidth: false,
                    pageLength: 100,
                    columns: [
                        {
                            data: 'EmployeeNameWithCodeRemarks',
                            render: function (employeeNameWithCodeRemarks) {
                                var splittedString = employeeNameWithCodeRemarks.split('$');
                                _id = parseInt(splittedString[3]); // id
                                if (splittedString[2] == 'true') {
                                    return `<span title='initial' data-name='${splittedString[0]}' style='color:red;'> ${splittedString[1]}</span> <input type='hidden' id='row_id_${_id}' value='${_id}'/>`;
                                } else {
                                    return `<span title='initial' data-name='${splittedString[0]}'> ${splittedString[1]}</span> <input type='hidden' id='row_id_${_id}' value='${_id}'/>`;
                                }

                            }
                        },
                        {
                            data: 'SectionName',
                            render: function (sectionName) {
                                return `<td title='initial'>${sectionName}</td>`;
                            }
                        },
                        {
                            data: 'DepartmentName',
                            render: function (departmentName) {
                                return `<td title='initial'>${departmentName}</td>`;
                            }
                        },
                        {
                            data: 'InchargeName',
                            render: function (inchargeName) {
                                return `<td title='initial'>${inchargeName}</td>`;
                            }
                        },
                        {
                            data: 'RoleName',
                            render: function (roleName) {
                                return `<td title='initial'>${roleName}</td>`;
                            }
                        },
                        {
                            data: 'ExplanationName',
                            render: function (explanationName) {
                                return `<span title='initial' class='forecast_explanation'>${explanationName}</span>`;
                            }
                        },
                        {
                            data: 'CompanyName',
                            render: function (companyName) {
                                _companyName = companyName;
                                return `<td title='initial'>${companyName}</td>`;
                            }
                        },
                        {
                            data: 'GradePoint',
                            render: function (gradePoint) {

                                if (_companyName.toLowerCase() == 'mw') {
                                    return `<td title='initial'>${gradePoint}</td>`;
                                }
                                else {
                                    return `<td title='initial'></td>`;
                                }

                            }
                        },
                        {
                            data: 'UnitPrice',
                            render: function (unitPrice) {
                                return `<span id='up_${_id}' data-unitprice=${unitPrice}> ${unitPrice}</span>`;
                            }
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='oct_${_id}'  onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)"  data-month='10' value='${_forecast[0].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='oct_${_id}'  onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)"  data-month='10' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='nov_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='11' value='${_forecast[1].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='nov_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='11' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='dec_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='12' value='${_forecast[2].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='dec_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='12' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jan_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='1' value='${_forecast[3].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='jan_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='1' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='feb_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='2' value='${_forecast[4].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='feb_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='2' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='mar_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='3' value='${_forecast[5].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='mar_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='3' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='apr_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='4' value='${_forecast[6].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='apr_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='4' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='may_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='5' value='${_forecast[7].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='may_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='5' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jun_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='6' value='${_forecast[8].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='jun_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='6' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jul_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='7' value='${_forecast[9].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='jul_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='7' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='aug_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='8' value='${_forecast[10].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='aug_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='8' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='sep_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='9' value='${_forecast[11].Points.toFixed(1)}' class='input_month'/></td>`;
                                } else {
                                    return `<td><input type='text' id='sep_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='9' value='0.0' class='input_month'/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='oct_output_${_id}' value='${_forecast[0].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='oct_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='nov_output_${_id}' value='${_forecast[1].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='nov_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='dec_output_${_id}' value='${_forecast[2].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='dec_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jan_output_${_id}' value='${_forecast[3].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='jan_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='feb_output_${_id}' value='${_forecast[4].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='feb_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='mar_output_${_id}' value='${_forecast[5].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='mar_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='apr_output_${_id}' value='${_forecast[6].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='apr_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='may_output_${_id}' value='${_forecast[7].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='may_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jun_output_${_id}' value='${_forecast[8].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='jun_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='jul_output_${_id}' value='${_forecast[9].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='jul_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='aug_output_${_id}' value='${_forecast[10].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='aug_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                        {
                            data: 'forecasts',
                            render: function (_forecast) {

                                if (_forecast.length > 0) {
                                    return `<td><input type='text' id='sep_output_${_id}' value='${_forecast[11].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                } else {
                                    return `<td><input type='text' id='sep_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                }

                            },
                            orderable: false
                        },
                    ],
                });

            },
            error: function () {
                $('#add_name_table_1 tbody').empty();
            }
        });
    }
}

function onCancel() {  
    LoaderShow();  
    LoadForecastData()    
}

function onSave() {   
    var saveFlag = false;
    const rows = document.querySelectorAll("#forecast_table > tbody > tr");
    if (rows.length <= 0) {
        alert('No Rows Selected');
        return false;
    }
    LoaderShow();
    $.each(rows, function (index, data) {
        var rowId = $(this).closest('tr').find('td').eq(0).children('input').val();
        var year = $('#period_search').find(":selected").val();
        var assignmentId = $('#row_id_' + rowId).val();

        var oct_point = $('#oct_' + rowId).val().replace(/,/g, '');
        var oct_output = $('#oct_output_' + rowId).val().replace(/,/g, '');

        var nov_point = $('#nov_' + rowId).val().replace(/,/g, '');
        var nov_output = $('#nov_output_' + rowId).val().replace(/,/g, '');


        var dec_point = $('#dec_' + rowId).val().replace(/,/g, '');
        var dec_output = $('#dec_output_' + rowId).val().replace(/,/g, '');

        var jan_point = $('#jan_' + rowId).val().replace(/,/g, '');
        var jan_output = $('#jan_output_' + rowId).val().replace(/,/g, '');

        var feb_point = $('#feb_' + rowId).val().replace(/,/g, '');
        var feb_output = $('#feb_output_' + rowId).val().replace(/,/g, '');

        var mar_point = $('#mar_' + rowId).val().replace(/,/g, '');
        var mar_output = $('#mar_output_' + rowId).val().replace(/,/g, '');

        var apr_point = $('#apr_' + rowId).val().replace(/,/g, '');
        var apr_output = $('#apr_output_' + rowId).val().replace(/,/g, '');

        var may_point = $('#may_' + rowId).val().replace(/,/g, '');
        var may_output = $('#may_output_' + rowId).val().replace(/,/g, '');

        var jun_point = $('#jun_' + rowId).val().replace(/,/g, '');
        var jun_output = $('#jun_output_' + rowId).val().replace(/,/g, '');

        var jul_point = $('#jul_' + rowId).val().replace(/,/g, '');
        var jul_output = $('#jul_output_' + rowId).val().replace(/,/g, '');

        var aug_point = $('#aug_' + rowId).val().replace(/,/g, '');
        var aug_output = $('#aug_output_' + rowId).val().replace(/,/g, '');

        var sep_point = $('#sep_' + rowId).val().replace(/,/g, '');
        var sep_output = $('#sep_output_' + rowId).val().replace(/,/g, '');

        var data = `10_${oct_point}_${oct_output},11_${nov_point}_${nov_output},12_${dec_point}_${dec_output},1_${jan_point}_${jan_output},2_${feb_point}_${feb_output},3_${mar_point}_${mar_output},4_${apr_point}_${apr_output},5_${may_point}_${may_output},6_${jun_point}_${jun_output},7_${jul_point}_${jul_output},8_${aug_point}_${aug_output},9_${sep_point}_${sep_output}`;

        $.ajax({
            url: '/api/Forecasts',
            type: 'GET',
            async: true,
            dataType: 'json',
            data: {                
                data: data,
                year: year,
                assignmentId: assignmentId
            },
            success: function (data) {                
                saveFlag = data;
                console.log("saved-1");
            },
            error: function (data) {
                $("#loading").css("display", "none");
                saveFlag = false;
                alert("Error please try again");
            }
        });
    });

    if (saveFlag) {
        console.log("saved-2");
        // $('#forecast_table').empty();
        // LoadForecastData();   
        Command: toastr["success"]('Data Saved Successfully', "Success")
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
    $('#forecast_table').empty();
    LoadForecastData();   
}
var expanded = false;
        $(document).on("click", function (event) {
            var $trigger = $(".forecast_multiselect");
            //alert("trigger: "+$trigger);
            if ($trigger !== event.target && !$trigger.has(event.target).length) {
                $(".commonselect").slideUp("fast");
                expanded = false;
            }
        });
        $(document).ready(function () {
            $('#forecast_name').click(function () {
                NameListSort("name_asc", "name_desc");
            });
            $('#forecast_section').click(function () {
                NameListSort("section_asc", "section_desc");
            });
            $('#forecast_department').click(function () {
                NameListSort("department_asc", "department_desc");
            });
            $('#forecast_incharge').click(function () {
                NameListSort("incharge_asc", "incharge_desc");
            });
            $('#forecast_role').click(function () {
                NameListSort("role_asc", "role_desc");
            });
            $('#forecast_explanation').click(function () {
                NameListSort("explanation_asc", "explanation_desc");
            });
            $('#forecast_company').click(function () {
                NameListSort("company_asc", "company_desc");
            });
            $('#forecast_grade').click(function () {
                NameListSort("grade_asc", "grade_desc");
            });
            $('#forecast_unitprice').click(function () {
                NameListSort("unit_asc", "unit_desc");
            });

            // multiple search by section
            $(document).on('click', '#sectionChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_sec_all") {
                    $(".section_checkbox").prop("checked", isSectionAllChk);
                } else {
                    $("#chk_sec_all").prop("checked", false);
                }


                clicked_checkbox = "";

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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var sectionSearchId = "";
                //if (!isSectionAllChk) {
                //    $.each(sectionCheckedBoxes, function(index, item) {
                //        if(sectionSearchId == ""){
                //            sectionSearchId = item.value;
                //        }
                //        else{
                //            sectionSearchId = sectionSearchId + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidSectionId").val(sectionSearchId);
            });

            // multiple search by departments
            $(document).on('click', '#departmentChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_dept_all") {
                    $(".department_checkbox").prop("checked", isDepartmentAllCheck);
                } else {
                    $("#chk_dept_all").prop("checked", false);
                }

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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var departmentIds = "";
                //$("#hidDepartmentId").val('');
                //if (!isDepartmentAllCheck) {
                //    $.each(departmentCheckedBoxes, function(index, item) {
                //        //sectionCheck.push(item.value);
                //        if(departmentIds == ""){
                //            departmentIds = item.value;
                //        }
                //        else{
                //            departmentIds = departmentIds + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidDepartmentId").val(departmentIds);
            });

            // multiple search by incharges
            $(document).on('click', '#inchargeChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_incharge_all") {
                    $(".incharge_checkbox").prop("checked", isInChargeAllCheck);
                } else {
                    $("#chk_incharge_all").prop("checked", false);
                }

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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var inchargeIds = "";
                //$("#hidInChargeId").val('');
                //if (!isInChargeAllCheck) {
                //    $.each(inchargeCheckedBoxes, function(index, item) {
                //        //sectionCheck.push(item.value);
                //        if(inchargeIds == ""){
                //            inchargeIds = item.value;
                //        }
                //        else{
                //            inchargeIds = inchargeIds + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidInChargeId").val(inchargeIds);
            });

            // multiple search by roles
            $(document).on('click', '#RoleChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_role_all") {
                    $(".role_checkbox").prop("checked", isRoleAllCheck);
                } else {
                    $("#chk_role_all").prop("checked", false);
                }


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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var roleIds = "";
                //$("#hidRoleId").val('');
                //if (!isRoleAllCheck) {
                //    $.each(roleCheckedBoxes, function(index, item) {
                //        //sectionCheck.push(item.value);
                //        if(roleIds == ""){
                //            roleIds = item.value;
                //        }
                //        else{
                //            roleIds = roleIds + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidRoleId").val(roleIds);

            });

            // multiple search by explanations
            $(document).on('click', '#ExplanationChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_explanation_all") {
                    $(".explanation_checkbox").prop("checked", isExplanationAllCheck);
                } else {
                    $("#chk_explanation_all").prop("checked", false);
                }

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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var explanationIds = "";
                //$("#hidExplanationId").val('');
                //if (!isExplanationAllCheck) {
                //    $.each(explanationCheckedBoxes, function(index, item) {
                //        //sectionCheck.push(item.value);
                //        if(explanationIds == ""){
                //            explanationIds = item.value;
                //        }
                //        else{
                //            explanationIds = explanationIds + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidExplanationId").val(explanationIds);
            });

            // multiple search by companies
            $(document).on('click', '#CompanyChks input[type="checkbox"]', function () {
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                var clicked_checkbox = $(this).attr("id")

                if (clicked_checkbox.toLowerCase() == "chk_comopany_all") {
                    $(".comopany_checkbox").prop("checked", isCompanytAllCheck);
                } else {
                    $("#chk_comopany_all").prop("checked", false);
                }

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

                $("#hidSectionId").val('');
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        sectionCheck.push(item.value);
                    });
                }
                $("#hidSectionId").val(sectionCheck);

                $("#hidDepartmentId").val('');
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        departmentCheck.push(item.value);
                    });
                }
                $("#hidDepartmentId").val(departmentCheck);

                $("#hidInChargeId").val('');
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        inchargeCheck.push(item.value);
                    });
                }
                $("#hidInChargeId").val(inchargeCheck);

                $("#hidRoleId").val('');
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        roleCheck.push(item.value);
                    });
                }
                $("#hidRoleId").val(roleCheck);

                $("#hidExplanationId").val('');
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        explanationCheck.push(item.value);
                    });
                }
                $("#hidExplanationId").val(explanationCheck);

                $("#hidCompanyid").val('');
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        companyCheck.push(item.value);
                    });
                }
                $("#hidCompanyid").val(companyCheck);

                //var companyIds = "";
                //$("#hidCompanyid").val('');
                //if (!isCompanytAllCheck) {
                //    $.each(companyCheckedBoxes, function(index, item) {
                //        //sectionCheck.push(item.value);
                //        if(companyIds == ""){
                //            companyIds = item.value;
                //        }
                //        else{
                //            companyIds = companyIds + ","+item.value;
                //        }
                //    });
                //}
                //$("#hidCompanyid").val(companyIds);
            });



            // var expanded = false;
            // $(document).on("click", function (event) {
            //     var $trigger = $(".forecast_multiselect");
            //     //alert("trigger: "+$trigger);
            //     if ($trigger !== event.target && !$trigger.has(event.target).length) {
            //         $(".commonselect").slideUp("fast");
            //     }
            // });

            ForecastSearchDropdownInLoad();
            // $('.container').blur(function (e) {
            //     $('.commonselect').fadeOut(100);
            // });

            $('#forecast_table').on('change', '.input_month', function () {

                //var rowId = parseInt($(this).closest('tr').data('rowid'));
                var rowId = $(this).closest('tr').find('td').eq(0).children('input').val();
                var month = parseInt($(this).data('month'));
                var unitPrice = parseFloat($('#up_' + rowId).data('unitprice').replace(/,/g, ''));
                var pointValue = parseFloat($(this).val());
                let result = 0;
                switch (month) {
                    case 1:
                        result = unitPrice * pointValue;
                        $('#jan_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 2:
                        result = unitPrice * pointValue;
                        $('#feb_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 3:
                        result = unitPrice * pointValue;
                        $('#mar_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 4:
                        result = unitPrice * pointValue;
                        $('#apr_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 5:
                        result = unitPrice * pointValue;
                        $('#may_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 6:
                        result = unitPrice * pointValue;
                        $('#jun_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 7:
                        result = unitPrice * pointValue;
                        $('#jul_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 8:
                        result = unitPrice * pointValue;
                        $('#aug_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 9:
                        result = unitPrice * pointValue;
                        $('#sep_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 10:
                        result = unitPrice * pointValue;
                        $('#oct_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 11:
                        result = unitPrice * pointValue;
                        $('#nov_output_' + rowId).val(result.toLocaleString());
                        break;
                    case 12:
                        result = unitPrice * pointValue;
                        $('#dec_output_' + rowId).val(result.toLocaleString());
                        break;

                }


            });


            $('#forecast_search_button').on('click', function () {
                var sectionId = $('#section_search').find(":selected").val();
                var inchargeId = $('#incharge_search').find(":selected").val();
                var departmentId = $('#department_search').find(":selected").val();
                var roleId = $('#role_search').find(":selected").val();
                var companyId = $('#company_search').find(":selected").val();
                var explanationId = $('#explanation_search').find(":selected").val();
                var employeeName = $('#identity_search').val();

                var year = $('#period_search').find(":selected").val();

                if (year == '' || year == undefined) {

                    alert('select year');
                    return false;
                }
                LoaderShow();

                $('#cancel_forecast').css('display', 'inline-block');
                $('#save_forecast').css('display', 'inline-block');

                //if (departmentId == undefined) {
                //    departmentId = '';
                //}
                let isSectionAllChk = $("#chk_sec_all").is(':checked');
                var isDepartmentAllCheck = $("#chk_dept_all").is(':checked');
                var isInChargeAllCheck = $("#chk_incharge_all").is(':checked');
                var isRoleAllCheck = $("#chk_role_all").is(':checked');
                var isExplanationAllCheck = $("#chk_explanation_all").is(':checked');
                var isCompanytAllCheck = $("#chk_comopany_all").is(':checked');

                //var employeeName = "";
                var sectionCheck = [];
                var departmentCheck = [];
                var inchargeCheck = [];
                var roleCheck = [];
                var explanationCheck = [];
                var companyCheck = [];

                var sectionCheckedBoxes = $('#sectionChks input[type="checkbox"]:checked');
                var departmentCheckedBoxes = $('#departmentChks input[type="checkbox"]:checked');
                var inchargeCheckedBoxes = $('#inchargeChks input[type="checkbox"]:checked');
                var roleCheckedBoxes = $('#RoleChks input[type="checkbox"]:checked');
                var explanationCheckedBoxes = $('#ExplanationChks input[type="checkbox"]:checked');
                var companyCheckedBoxes = $('#CompanyChks input[type="checkbox"]:checked');

                sectionId = "";
                if (!isSectionAllChk) {
                    $.each(sectionCheckedBoxes, function (index, item) {
                        //sectionCheck.push(item.value);
                        if (sectionId == "") {
                            sectionId = item.value;
                        } else {
                            sectionId = sectionId + "##" + item.value;
                        }
                    });
                }

                departmentId = "";
                if (!isDepartmentAllCheck) {
                    $.each(departmentCheckedBoxes, function (index, item) {
                        //departmentCheck.push(item.value);
                        if (departmentId == "") {
                            departmentId = item.value;
                        } else {
                            departmentId = departmentId + "##" + item.value;
                        }
                    });
                }

                inchargeId = "";
                if (!isInChargeAllCheck) {
                    $.each(inchargeCheckedBoxes, function (index, item) {
                        //inchargeCheck.push(item.value);
                        if (inchargeId == "") {
                            inchargeId = item.value;
                        } else {
                            inchargeId = inchargeId + "##" + item.value;
                        }
                    });
                }
                roleId = "";
                if (!isRoleAllCheck) {
                    $.each(roleCheckedBoxes, function (index, item) {
                        //roleCheck.push(item.value);
                        if (roleId == "") {
                            roleId = item.value;
                        } else {
                            roleId = roleId + "##" + item.value;
                        }
                    });
                }
                explanationId = "";
                if (!isExplanationAllCheck) {
                    $.each(explanationCheckedBoxes, function (index, item) {
                        //explanationCheck.push(item.value);
                        if (explanationId == "") {
                            explanationId = item.value;
                        } else {
                            explanationId = explanationId + "##" + item.value;
                        }
                    });
                }
                companyId = "";
                if (!isCompanytAllCheck) {
                    $.each(companyCheckedBoxes, function (index, item) {
                        //companyCheck.push(item.value);
                        if (companyId == "") {
                            companyId = item.value;
                        } else {
                            companyId = companyId + "##" + item.value;
                        }
                    });
                }

                var data_info = {
                    employeeName: employeeName,
                    sectionId: sectionId,
                    departmentId: departmentId,
                    inchargeId: inchargeId,
                    roleId: roleId,
                    explanationId: explanationId,
                    companyId: companyId,
                    status: '', year: ''
                };

                globalSearchObject = data_info;

                $.ajax({
                    //url: `/api/utilities/SearchForecastEmployee`,
                    //type: 'GET',
                    //dataType: 'json',
                    //data: data_info,
                    url: `/api/utilities/SearchForecastEmployee`,
                    contentType: 'application/json',
                    type: 'GET',
                    dataType: 'json',
                    data: data_info,
                    success: function (data) {
                        LoaderHide();

                        $('#forecast_table>thead').empty();
                        $('#forecast_table>tbody').empty();
                        $('#forecast_table>thead').append(`

                                                <tr>
                                                    <th id="forecast_name">Name </th>
                                                    <th id="forecast_section">Section </th>
                                                    <th id="forecast_department">Department </th>
                                                    <th id="forecast_incharge">In-Charge </th>
                                                    <th id="forecast_role">Role </th>
                                                    <th id="forecast_explanation">Explanation </th>
                                                    <th id="forecast_company">Company Name </i></th>
                                                    <th id="forecast_grade">Grade </th>
                                                    <th id="forecast_unitprice"><span>Unit Price</span> </th>
                                                    <th>10月</th>
                                                    <th>11月</th>
                                                    <th>12月</th>
                                                    <th>1月</th>
                                                    <th>2月</th>
                                                    <th>3月</th>
                                                    <th>4月</th>
                                                    <th>5月</th>
                                                    <th>6月</th>
                                                    <th>7月</th>
                                                    <th>8月</th>
                                                    <th>9月</th>
                                                    <th>Oct</th>
                                                    <th>Nov</th>
                                                    <th>Dec</th>
                                                    <th>Jan</th>
                                                    <th>Feb</th>
                                                    <th>Mar</th>
                                                    <th>Apr</th>
                                                    <th>May</th>
                                                    <th>Jun</th>
                                                    <th>Jul</th>
                                                    <th>Aug</th>
                                                    <th>Sep</th>
                                                </tr>
                                            `);
                        //Forecast_DatatableLoad(data);
                        var _id = 0;
                        var _companyName = '';
                        var _oct = 0;
                        var _octTotal = 0;
                        $('#forecast_table').DataTable({
                            destroy: true,
                            data: data,
                            searching: false,
                            bLengthChange: false,
                            autoWidth: false,
                            pageLength: 100,
                            columns: [
                                //{
                                //    data: 'Id',
                                //    render: function (id) {
                                //        _id = id;
                                //        return null;
                                //    },
                                //    //ordering: false

                                //},
                                {
                                    data: 'EmployeeNameWithCodeRemarks',
                                    render: function (employeeNameWithCodeRemarks) {
                                        var splittedString = employeeNameWithCodeRemarks.split('$');
                                        _id = parseInt(splittedString[3]); // id
                                        if (splittedString[2] == 'true') {
                                            return `<span title='initial' data-name='${splittedString[0]}' style='color:red;'> ${splittedString[1]}</span> <input type='hidden' id='row_id_${_id}' value='${_id}'/>`;
                                        } else {
                                            return `<span title='initial' data-name='${splittedString[0]}'> ${splittedString[1]}</span> <input type='hidden' id='row_id_${_id}' value='${_id}'/>`;
                                        }

                                    }
                                },
                                {
                                    data: 'SectionName',
                                    render: function (sectionName) {
                                        return `<td title='initial'>${sectionName}</td>`;
                                    }
                                },
                                {
                                    data: 'DepartmentName',
                                    render: function (departmentName) {
                                        return `<td title='initial'>${departmentName}</td>`;
                                    }
                                },
                                {
                                    data: 'InchargeName',
                                    render: function (inchargeName) {
                                        return `<td title='initial'>${inchargeName}</td>`;
                                    }
                                },
                                {
                                    data: 'RoleName',
                                    render: function (roleName) {
                                        return `<td title='initial'>${roleName}</td>`;
                                    }
                                },
                                {
                                    data: 'ExplanationName',
                                    render: function (explanationName) {
                                        return `<span title='initial' class='forecast_explanation'>${explanationName}</span>`;
                                    }
                                },
                                {
                                    data: 'CompanyName',
                                    render: function (companyName) {
                                        _companyName = companyName;
                                        return `<td title='initial'>${companyName}</td>`;
                                    }
                                },
                                {
                                    data: 'GradePoint',
                                    render: function (gradePoint) {

                                        if (_companyName.toLowerCase() == 'mw') {
                                            return `<td title='initial'>${gradePoint}</td>`;
                                        }
                                        else {
                                            return `<td title='initial'></td>`;
                                        }

                                    }
                                },
                                {
                                    data: 'UnitPrice',
                                    render: function (unitPrice) {
                                        return `<span id='up_${_id}' data-unitprice=${unitPrice}> ${unitPrice}</span>`;
                                    }
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='oct_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)"  data-month='10' value='${_forecast[0].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='oct_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)"  data-month='10' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='nov_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='11' value='${_forecast[1].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='nov_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='11' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='dec_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='12' value='${_forecast[2].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='dec_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='12' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jan_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='1' value='${_forecast[3].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jan_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='1' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='feb_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='2' value='${_forecast[4].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='feb_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='2' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='mar_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='3' value='${_forecast[5].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='mar_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='3' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='apr_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='4' value='${_forecast[6].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='apr_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='4' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='may_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='5' value='${_forecast[7].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='may_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='5' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jun_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='6' value='${_forecast[8].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jun_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='6' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jul_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='7' value='${_forecast[9].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jul_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='7' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='aug_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='8' value='${_forecast[10].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='aug_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='8' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='sep_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='9' value='${_forecast[11].Points.toFixed(1)}' class='input_month'/></td>`;
                                        } else {
                                            return `<td><input type='text' id='sep_${_id}' onfocus ="checkPoint_onmouseover(this);" onclick="checkPoint_click(this)" onChange="checkPoint(this)" data-month='9' value='0.0' class='input_month'/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='oct_output_${_id}' value='${_forecast[0].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='oct_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='nov_output_${_id}' value='${_forecast[1].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='nov_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='dec_output_${_id}' value='${_forecast[2].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='dec_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jan_output_${_id}' value='${_forecast[3].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jan_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='feb_output_${_id}' value='${_forecast[4].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='feb_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='mar_output_${_id}' value='${_forecast[5].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='mar_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='apr_output_${_id}' value='${_forecast[6].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='apr_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='may_output_${_id}' value='${_forecast[7].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='may_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jun_output_${_id}' value='${_forecast[8].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jun_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='jul_output_${_id}' value='${_forecast[9].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='jul_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='aug_output_${_id}' value='${_forecast[10].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='aug_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                                {
                                    data: 'forecasts',
                                    render: function (_forecast) {

                                        if (_forecast.length > 0) {
                                            return `<td><input type='text' id='sep_output_${_id}' value='${_forecast[11].Total}' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        } else {
                                            return `<td><input type='text' id='sep_output_${_id}' value='0' style='background-color:#E6E6E3;text-align:right;' readonly/></td>`;
                                        }

                                    },
                                    orderable: false
                                },
                            ],
                        });

                    },
                    error: function () {
                        $('#add_name_table_1 tbody').empty();
                    }
                });


            });

            $(document).on('change', '#section_search', function () {

                var sectionId = $(this).val();

                $.getJSON(`/api/utilities/DepartmentsBySection/${sectionId}`)
                    .done(function (data) {
                        $('#department_search').empty();
                        $('#department_search').append(`<option value=''>Select Department</option>`);
                        $.each(data, function (key, item) {
                            $('#department_search').append(`<option value='${item.Id}'>${item.DepartmentName}</option>`);
                        });
                    });
            });


        });