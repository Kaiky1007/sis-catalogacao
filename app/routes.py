import os
from flask import render_template, request, redirect, url_for, current_app
from werkzeug.utils import secure_filename
from datetime import datetime
from app import app, db
from .models import Ficha

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def processar_grupo_checkbox(prefixo, lista_opcoes, form_data):
    resultado = {}
    for opcao in lista_opcoes:
        chave_html = f"{prefixo}_{opcao}"
        resultado[opcao] = True if form_data.get(chave_html) else False
    
    campo_outro = form_data.get(f"{prefixo}_outro")
    if campo_outro:
        resultado['outro_texto'] = campo_outro
        
    return resultado

@app.route('/')
def index():
    fichas = Ficha.query.order_by(Ficha.id.desc()).all()
    return render_template('index.html', fichas=fichas)

@app.route('/nova')
def nova_ficha():
    return render_template('fichas.html', ficha=None)

@app.route('/criar', methods=['POST'])
def criar_ficha():
    try:
        caminho_relativo = None
        if 'foto' in request.files:
            file = request.files['foto']
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
                
                upload_folder = os.path.join(current_app.root_path, 'static', 'uploads')
                os.makedirs(upload_folder, exist_ok=True)
                
                file.save(os.path.join(upload_folder, filename))
                caminho_relativo = f"uploads/{filename}"

        data_str = request.form.get('data_preenchimento')
        data_obj = None
        if data_str:
            data_obj = datetime.strptime(data_str, '%Y-%m-%d').date()
        else:
            data_obj = datetime.now().date()

        nova_ficha = Ficha(
            numero_ficha=request.form.get('numero_ficha'),
            avaliacao=request.form.get('avaliacao'),
            autor=request.form.get('autor'),
            titulo=request.form.get('titulo'),
            registro=request.form.get('registro'),
            n_chamada=request.form.get('n_chamada'),
            secao_guarda=request.form.get('secao_guarda'),
            data_obra=request.form.get('data_obra'),
            paginas=request.form.get('paginas'),
            dimensoes=request.form.get('dimensoes'),
            observacoes=request.form.get('observacoes'),
            tecnico_nome=request.form.get('tecnico_nome'),
            data_preenchimento=data_obj,
            foto_path=caminho_relativo,

            especificacao_material=processar_grupo_checkbox('material', 
                ['album', 'folheto', 'manuscrito', 'planta', 'brochura', 'gravura', 
                 'mapa', 'pergaminho', 'certificado', 'impresso', 'partitura', 'desenho', 'livro', 'periodico'], 
                request.form),

            tipo_suporte=processar_grupo_checkbox('suporte', 
                ['couche', 'jornal', 'feito_mao', 'madeira'], 
                request.form),

            estado_conservacao=processar_grupo_checkbox('estado', 
                ['encadernada', 'sem_encadernacao', 'inteira', 'meia_com_cantos', 'meia_sem_cantos'], 
                request.form),

            deterioracoes=processar_grupo_checkbox('det', 
                ['abrasao', 'costura_fragil', 'mancha', 'rompimento', 'arranhao', 
                 'descoloracao', 'perda_lombada', 'sujidades'], 
                request.form),

            tratamento_planos=processar_grupo_checkbox('trat_plano', 
                ['diagnostico', 'retirada_sujidades', 'trincha', 'higienizacao', 'retirada_fitas', 
                 'po_borracha', 'desacidificacao', 'arrefecimento', 'reestruturacao', 'remendos', 
                 'enxertos', 'velaturas', 'planificacao', 'acondicionamento', 'portfolio', 
                 'passe_partout', 'pasta', 'envelope', 'jaqueta'], 
                request.form),
                
            tratamento_volumes=processar_grupo_checkbox('trat_vol', 
                ['fumigacao', 'fungos', 'insetos', 'higienizacao', 'trincha', 'reestruturacao', 
                 'lombada', 'lombada_capa', 'folhas', 'encadernacao', 'inteira', 'meia_sem_cantos', 
                 'costura', 'douracao', 'punho', 'maquina', 'acondicionamento', 'caixa_cruz', 'caixa_cadarco'], 
                request.form)
        )

        db.session.add(nova_ficha)
        db.session.commit()
        
        return redirect(url_for('ver_ficha', id=nova_ficha.id))

    except Exception as e:
        db.session.rollback()
        print(f"Erro ao salvar: {e}")
        return "Erro ao salvar a ficha. Verifique o console.", 500

@app.route('/ficha/<int:id>')
def ver_ficha(id):
    ficha = Ficha.query.get_or_404(id)
    return render_template('fichas.html', ficha=ficha)
