document.addEventListener('DOMContentLoaded', function() {
    
    const inputFoto = document.querySelector('input[name="foto"]');
    const imgPreview = document.querySelector('.img-preview');
    const containerFoto = imgPreview ? imgPreview.parentElement : null;

    if (inputFoto) {
        inputFoto.addEventListener('change', function(event) {
            const arquivo = event.target.files[0];

            if (arquivo) {
                const leitor = new FileReader();

                leitor.onload = function(e) {
                    if (imgPreview) {
                        imgPreview.src = e.target.result;
                        imgPreview.style.display = 'block';
                    } else {
                        if (containerFoto) {
                            const textoAntigo = containerFoto.querySelector('div');
                            if(textoAntigo) textoAntigo.remove();

                            const novaImg = document.createElement('img');
                            novaImg.src = e.target.result;
                            novaImg.classList.add('img-preview');
                            containerFoto.insertBefore(novaImg, inputFoto);
                        }
                    }
                }

                leitor.readAsDataURL(arquivo);
            }
        });
    }

    const checkEncadernada = document.querySelector('input[name="estado_encadernada"]');
    const checkSemEncadernacao = document.querySelector('input[name="estado_sem_encadernacao"]');

    if (checkEncadernada && checkSemEncadernacao) {
        checkEncadernada.addEventListener('change', function() {
            if (this.checked) {
                checkSemEncadernacao.checked = false;
            }
        });

        checkSemEncadernacao.addEventListener('change', function() {
            if (this.checked) {
                checkEncadernada.checked = false;
                const tiposEncadernacao = ['estado_inteira', 'estado_meia_com_cantos', 'estado_meia_sem_cantos'];
                tiposEncadernacao.forEach(nome => {
                    const el = document.querySelector(`input[name="${nome}"]`);
                    if(el) el.checked = false;
                });
            }
        });
    }

    const alertas = document.querySelectorAll('.alert');
    if (alertas.length > 0) {
        setTimeout(function() {
            alertas.forEach(function(alerta) {
                alerta.style.transition = "opacity 1s";
                alerta.style.opacity = "0";
                
                setTimeout(() => alerta.remove(), 1000);
            });
        }, 4000);
    }

});
