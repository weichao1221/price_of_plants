
var table = document.getElementById("table");
var rows = table.getElementsByTagName("tr");

var rowsPerPage = 10;

var pageCount = Math.ceil(rows / rowsPerPage);

function showPage (page) {
    var startIndex = (page - 1) * rowsPerPage;
    var endIndex = page + rowsPerPage;

    for (var i = 0; i < rows.length; i++) {
        if (i >= startIndex && i < endIndex) {
            rows[i].style.display = '';
        } else {
            rows[i].style.display = 'none';
        }
    }
}

// 默认显示第一页
showPage(1);

// 添加分页按钮

var paginationDiv = document.createElement("div");
paginationDiv.id = "paginationDiv";

for (var i = 1; i <= pageCount; i++) {
    var button = document.createElement('button');
    button.textContent = i;
    button.addEventListener('click', function() {
        showPage(parseInt(this.textContent));
    });
    paginationDiv.appendChild(button);
}

// 将分页按钮添加到页面
document.body.appendChild(paginationDiv);