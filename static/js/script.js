document.getElementById('submitBtn').addEventListener('click', function() {
    var formData = new FormData();
    formData.append('fileUpload', document.getElementById('fileUpload').files[0]);

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/upload', true);

    xhr.onload = function () {
        if (xhr.status === 200) {
            // Sucesso
            document.getElementById('message').innerText = 'Arquivo enviado com sucesso!';
        } else {
            // Erro
            document.getElementById('message').innerText = 'Falha no envio do arquivo.';
        }
    };

    xhr.send(formData);
});
