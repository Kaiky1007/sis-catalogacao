import os
import pandas as pd
from flask import render_template, request, redirect, url_for, current_app
from werkzeug.utils import secure_filename
from datetime import datetime
from app import app, db
from .models import Ficha
from flask import flash

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

@app.route('/acervo')
def listar_acervo():
    fichas = Ficha.query.order_by(Ficha.id.desc()).all()
    return render_template('lista.html', fichas=fichas) 

def converter_booleano(valor):
    if pd.isna(valor) or valor == '':
        return False
    val_str = str(valor).lower().strip()
    return val_str in ['sim', 's', 'true', '1', 'x', 'yes', 'checked']

def converter_avaliacao(valor):
    if pd.isna(valor):
        return 2
    val_str = str(valor).lower().strip()
    if 'bom' in val_str:
        return 1
    if 'mau' in val_str or 'ruim' in val_str:
        return 3
    return 2

@app.route('/importar', methods=['POST'])
def importar_planilha():
    if 'arquivo_excel' not in request.files:
        flash('Nenhum arquivo enviado.', 'error')
        return redirect(url_for('listar_acervo'))
    
    arquivo = request.files['arquivo_excel']
    if arquivo.filename == '':
        flash('Nenhum arquivo selecionado.', 'error')
        return redirect(url_for('listar_acervo'))

    try:
        if arquivo.filename.endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python')
        else:
            df = pd.read_excel(arquivo)
        
        imported_count = 0
        
        for index, row in df.iterrows():
            num_ficha = str(row.get('numero_ficha', ''))
            if not num_ficha or Ficha.query.filter_by(numero_ficha=num_ficha).first():
                continue 

            data_str = str(row.get('data', ''))
            try:
                data_final = pd.to_datetime(data_str, dayfirst=True).date()
            except:
                data_final = datetime.now().date()

            nova_ficha = Ficha(
                numero_ficha=num_ficha,
                avaliacao=converter_avaliacao(row.get('avaliacao')),
                autor=row.get('autor', ''),
                titulo=row.get('titulo', ''),
                registro=str(row.get('registro', '')),
                n_chamada=str(row.get('n_chamada', '')),
                secao_guarda=row.get('secao_guarda', ''),
                data_obra=str(row.get('data_obra', '')),
                paginas=str(row.get('paginas', '')),
                dimensoes=str(row.get('dimensoes', '')),
                observacoes=row.get('observacoes', ''),
                tecnico_nome=row.get('tecnico', ''),
                foto_path=row.get('caminho_foto', ''),
                data_preenchimento=data_final,
                especificacao_material={
                    'album': converter_booleano(row.get('album')),
                    'folheto': converter_booleano(row.get('folheto')),
                    'manuscrito': converter_booleano(row.get('manuscrito')),
                    'planta': converter_booleano(row.get('planta')),
                    'brochura': converter_booleano(row.get('brochura')),
                    'gravura': converter_booleano(row.get('gravura')),
                    'mapa': converter_booleano(row.get('mapa')),
                    'pergaminho': converter_booleano(row.get('pergaminho')),
                    'certificado': converter_booleano(row.get('certificado')),
                    'impresso': converter_booleano(row.get('impresso')),
                    'partitura': converter_booleano(row.get('partitura')),
                    'desenho': converter_booleano(row.get('desenho')),
                    'livro': converter_booleano(row.get('livro')),
                    'periodico': converter_booleano(row.get('periodico')),
                    'outro_texto': row.get('material_outro', '')
                },
                tipo_suporte={
                    'couche': converter_booleano(row.get('papel_couche')),
                    'jornal': converter_booleano(row.get('papel_jornal')),
                    'feito_mao': converter_booleano(row.get('papel_feito_mao')),
                    'madeira': converter_booleano(row.get('papel_madeira')),
                    'outro_texto': row.get('suporte_outro', '')
                },
                estado_conservacao={
                    'encadernada': converter_booleano(row.get('obra_encadernada')),
                    'sem_encadernacao': converter_booleano(row.get('obra_sem_encadernacao')),
                    'inteira': converter_booleano(row.get('encadernacao_inteira')),
                    'meia_com_cantos': converter_booleano(row.get('meia_com_cantos')),
                    'meia_sem_cantos': converter_booleano(row.get('meia_sem_cantos'))
                },
                deterioracoes={
                    'abrasao': converter_booleano(row.get('abrasao')),
                    'costura_fragil': converter_booleano(row.get('costura_fragil')),
                    'mancha': converter_booleano(row.get('mancha')),
                    'rompimento': converter_booleano(row.get('rompimento')),
                    'arranhao': converter_booleano(row.get('arranhao')),
                    'descoloracao': converter_booleano(row.get('descoloracao')),
                    'perda_lombada': converter_booleano(row.get('perda_lombada')),
                    'sujidades': converter_booleano(row.get('sujidades'))
                },
                tratamento_planos={
                    'diagnostico': converter_booleano(row.get('diagnostico')),
                    'retirada_sujidades': converter_booleano(row.get('retirada_sujidades')),
                    'trincha': converter_booleano(row.get('trincha_macia')),
                    'higienizacao': converter_booleano(row.get('higienizacao')),
                    'retirada_fitas': converter_booleano(row.get('retirada_fitas')),
                    'po_borracha': converter_booleano(row.get('po_borracha')),
                    'desacidificacao': converter_booleano(row.get('desacidificacao')),
                    'arrefecimento': converter_booleano(row.get('arrefecimento')),
                    'reestruturacao': converter_booleano(row.get('reestruturacao')),
                    'remendos': converter_booleano(row.get('remendos')),
                    'enxertos': converter_booleano(row.get('enxertos')),
                    'velaturas': converter_booleano(row.get('velaturas')),
                    'planificacao': converter_booleano(row.get('planificacao')),
                    'acondicionamento': converter_booleano(row.get('acondicionamento')),
                    'portfolio': converter_booleano(row.get('portfolio')),
                    'passe_partout': converter_booleano(row.get('passe_partout')),
                    'pasta': converter_booleano(row.get('pasta')),
                    'envelope': converter_booleano(row.get('envelope')),
                    'jaqueta': converter_booleano(row.get('jaqueta')),
                    'outro_texto': row.get('tratamento_plano_outro', '')
                },
                tratamento_volumes={
                    'fumigacao': converter_booleano(row.get('fumigacao')),
                    'fungos': converter_booleano(row.get('fungos')),
                    'insetos': converter_booleano(row.get('insetos')),
                    'higienizacao': converter_booleano(row.get('higienizacao_volumes')),
                    'trincha': converter_booleano(row.get('trincha_macia_volumes')),
                    'reestruturacao': converter_booleano(row.get('reestruturacao_volumes')),
                    'lombada': converter_booleano(row.get('lombada')),
                    'lombada_capa': converter_booleano(row.get('lombada_capa')),
                    'folhas': converter_booleano(row.get('folhas')),
                    'encadernacao': converter_booleano(row.get('encadernacao')),
                    'inteira': converter_booleano(row.get('inteira_volumes')),
                    'meia_sem_cantos': converter_booleano(row.get('meia_sem_cantos_volumes')),
                    'costura': converter_booleano(row.get('costura')),
                    'douracao': converter_booleano(row.get('douracao')),
                    'punho': converter_booleano(row.get('a_punho')),
                    'maquina': converter_booleano(row.get('a_maquina')),
                    'acondicionamento': converter_booleano(row.get('acondicionamento_volumes')),
                    'caixa_cruz': converter_booleano(row.get('caixa_cruz')),
                    'caixa_cadarco': converter_booleano(row.get('caixa_cadarco')),
                    'outro_texto': row.get('tratamento_volumes_outro', '')
                }
            )

            db.session.add(nova_ficha)
            imported_count += 1
        
        db.session.commit()
        flash(f'Sucesso! {imported_count} fichas importadas.', 'success')
        return redirect(url_for('listar_acervo'))

    except Exception as e:
        db.session.rollback()
        print(f"Erro IMPORTACAO: {e}") 
        flash('Erro ao processar arquivo. Verifique se as colunas estão corretas.', 'danger')
        return redirect(url_for('listar_acervo'))
