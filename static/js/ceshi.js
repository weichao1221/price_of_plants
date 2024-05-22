function button_click() {
    const text_1 = document.getElementById('text_1');
    const text_2 = document.getElementById('text_2');
    const text_3 = document.getElementById('text_3');
    if (text_1.value === "" || text_2.value === "" || text_3.value === "") {
        alert("不允许为空值！");
    } else {
        console.log(text_1.value, text_2.value, text_2.value);
    }
}

function clear_content() {
    var text_1 = document.getElementById('text_1');
    var text_2 = document.getElementById('text_2');
    var text_3 = document.getElementById('text_3');
    // console.log(text_1.value);
    text_1.value = "";
    text_2.value = "";
    text_3.value = "";
    console.log("输入框的内容已经清除！");
}

document.getElementById("menuToggle").addEventListener('click', function () {
    var leftMenu = document.getElementById("leftMenu");
    leftMenu.classList.toggle("open");
});