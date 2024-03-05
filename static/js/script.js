document.getElementById('submitBtn').addEventListener('click', function() {
    document.getElementById('loading').style.display = 'block';

    var formData = new FormData();
    formData.append('fileUpload', document.getElementById('hiddenFileUpload').files[0]);

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/upload', true);

    xhr.onload = function () {
        document.getElementById('loading').style.display = 'none';

        if (xhr.status === 200) {
            document.getElementById('modal').style.display = 'block';
        } else {
            document.getElementById('modal').style.display = 'none';
        }
    };

    xhr.onerror = function () {
        document.getElementById('loading').style.display = 'none';
        document.getElementById('modal').style.display = 'none';
    };

    xhr.send(formData);
});

document.addEventListener('DOMContentLoaded', function() {
    const customBtn = document.getElementById('fileUpload');
    const realFileBtn = document.getElementById('hiddenFileUpload');
    const customTxt = document.getElementById('fileChosen');

    if (customBtn) {
        customBtn.addEventListener('click', function() {
            realFileBtn.click();
        });
    }

    if (realFileBtn) {
        realFileBtn.addEventListener('change', function() {
            if (realFileBtn.value) {
                customTxt.innerHTML = realFileBtn.value.match(/[\/\\]([\w\d\s\.\-\(\)]+)$/)[1]; // Extrai apenas o nome do arquivo
            } else {
                customTxt.innerHTML = 'Nenhum arquivo escolhido';
            }
        });
    }
});

document.querySelector('.close-button').addEventListener('click', function() {
    document.getElementById('modal').style.display = 'none';
});

