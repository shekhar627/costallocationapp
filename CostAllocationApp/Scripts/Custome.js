function CreateMultipleInput() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='division-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a section code.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-4'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>Section Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='division-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a section name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";

    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='CreateMultipleInput();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div').append(multipleHtmlInputs);
}

function Department_MultiRow() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>Department Code<span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a department code  .";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>Department Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a department name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <label for='section-name'>Section Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "        <select id='section-name' class='form-control col-12' style='width:100% !important; border:1px solid #e3e3e3 !important;'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Select Section</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-1</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-2</option>";
    multipleHtmlInputs = multipleHtmlInputs + "            <option value=''>Section-3</option>";
    multipleHtmlInputs = multipleHtmlInputs + "        </select>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "            Please provide a department name.";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='Department_MultiRow();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div-department').append(multipleHtmlInputs);
}
function CreateMultipleInputInCharge() {
    var multipleHtmlInputs = "";
    multipleHtmlInputs = multipleHtmlInputs + "<div class='row'> ";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='divison-name'>In-Charge Code<span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='divison-name' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a In-Charge code  .";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-4'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code'>In-Charge Name <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <input class='form-control' id='section-code' type='text'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <div class='department-name-error'>";
    multipleHtmlInputs = multipleHtmlInputs + "                Please provide a In-Charge name.";
    multipleHtmlInputs = multipleHtmlInputs + "            </div>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    <div class='col-sm-3'>";
    multipleHtmlInputs = multipleHtmlInputs + "        <div class='form-group'>";
    multipleHtmlInputs = multipleHtmlInputs + "            <label for='section-code' id='plus-icon'>Section Code <span class='text-danger'>*</span></label>";
    multipleHtmlInputs = multipleHtmlInputs + "            <button type='button' class='form-control' onclick='CreateMultipleInput();' style='border:none;padding-right:139px;'><i class='fa fa-plus'></i></button>";
    multipleHtmlInputs = multipleHtmlInputs + "        </div>";
    multipleHtmlInputs = multipleHtmlInputs + "    </div>";
    multipleHtmlInputs = multipleHtmlInputs + "</div>";

    $('#main-div-in-charge').append(multipleHtmlInputs);
}



var expanded = false;

function SectionCheck() {
    var checkboxes = document.getElementById("sectionChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function DepartmentCheck() {
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
    var checkboxes = document.getElementById("RoleChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}
function CompanyCheck() {
    var checkboxes = document.getElementById("CompanyChks");
    if (!expanded) {
        checkboxes.style.display = "block";
        expanded = true;
    } else {
        checkboxes.style.display = "none";
        expanded = false;
    }
}

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
