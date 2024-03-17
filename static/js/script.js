function closeModal() {
    $('#modal').hide();
    $('#hiddenFileUpload').val('');
    $('#fileChosen').text('Nenhum arquivo escolhido');
    $('#submitBtn').prop('disabled', true);
}

$(document).ready(function() {
    $('#submitBtn').click(function() {
        $('#loading').show();

        var formData = new FormData();
        formData.append('fileUpload', $('#hiddenFileUpload')[0].files[0]);

        $.ajax({
            url: '/upload',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function() {
                $('#loading').hide();
                $('#modal').toggle(true);
            },
            error: function() {
                $('#loading').hide();
                $('#modal').hide();
            }
        });
    });

    $('#fileUpload').click(function() {
        $('#hiddenFileUpload').click();
    });

    $('#hiddenFileUpload').change(function() {
        var fileName = $(this).val().split("\\").pop();
        $('#submitBtn').prop('disabled', !fileName);
        $('#fileChosen').text(fileName ? fileName : 'Nenhum arquivo escolhido');
    });



    $('.close-button').click(closeModal);
    $('#uploadZipButton').click(closeModal);
});
