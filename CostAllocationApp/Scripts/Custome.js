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

/*------------------------------------Reporting-----------------------------------------*/

$(document).ready(function () {
    // ------------------Pie-Chart---------------------------//
    google.charts.load('current', { 'packages': ['corechart'] });
    google.charts.setOnLoadCallback(CostManagement);
    google.charts.setOnLoadCallback(Employees);
    google.charts.setOnLoadCallback(Grades);

    // Draw the chart and set the chart values
    function CostManagement() {
        var data = google.visualization.arrayToDataTable([
            ['Task', 'Number of Modules'],
            ['Section', 5],
            ['Department', 10],
            ['Companies', 20],
            ['Fiscal year calculated', 6],
            ['Grades', 18]
        ]);

        // Optional; add a title and set the width and height of the chart
        var options = { 'title': 'Cost Management System', 'width': 550, 'height': 400 };

        // Display the chart inside the <div> element with id="piechart"
        var chart = new google.visualization.PieChart(document.getElementById('cost_management'));
        chart.draw(data, options);
    }

    function Employees() {
        var data = google.visualization.arrayToDataTable([
            ['Task', 'Employees'],
            ['Employess', 100],
            ['Employee assign to Section', 85],
            ['Employee assign to Department', 90],
            ['Employees with company', 95],
            ['Employees Grade', 50]
        ]);

        // Optional; add a title and set the width and height of the chart
        var options = { 'title': 'Employees', 'width': 550, 'height': 400 };

        // Display the chart inside the <div> element with id="piechart"
        var chart = new google.visualization.PieChart(document.getElementById('employees'));
        chart.draw(data, options);
    }
    function Grades() {
        var data = google.visualization.arrayToDataTable([
            ['Task', 'Grades'],
            ['Number Of Grades',55],
            ['Number of MW companies', 10],
            ['Other Companies', 45]
        ]);

        // Optional; add a title and set the width and height of the chart
        var options = { 'title': 'Grade System', 'width': 550, 'height': 400 };

        // Display the chart inside the <div> element with id="piechart"
        var chart = new google.visualization.PieChart(document.getElementById('grade_system'));
        chart.draw(data, options);
    }




    //------------------------------- Bar-Chart----------------------//
    var xValues = ["Year: 2023 ", "Year: 2022", "Year: 2021", "Year: 2020", "Year:2019"];
    var yValues = [65, 55, 50, 25, 11];
    var barColors = ["red", "green", "blue", "orange", "brown"];

    new Chart("myChart", {
        type: "bar",
        data: {
            labels: xValues,
            datasets: [{
                backgroundColor: barColors,
                data: yValues
            }]
        },
        options: {
            legend: { display: false },
            title: {
                display: true,
                text: "Forecast through the year"
            }
        }
    });
});