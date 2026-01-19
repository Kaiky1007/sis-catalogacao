import os
import pandas as pd
from flask import render_template, request, redirect, url_for, current_app, flash
from werkzeug.utils import secure_filename
from datetime import datetime
from app import app, db
from .models import Ficha
import io
from flask import send_file

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
    if val_str.endswith('.0'):
        val_str = val_str.replace('.0', '')
    valores_true = ['sim', 's', 'true', '1', 'x', 'yes', 'checked', 'verdadeiro', 'on']
    return val_str in valores_true

def converter_avaliacao(valor):
    if pd.isna(valor): 
        return 2
    val_str = str(valor).lower().strip()
    if val_str in ['1', 'bom']: 
        return 1
    if val_str in ['3', 'mau', 'ruim']: 
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
        
        df.columns = df.columns.str.strip()

        imported_count = 0
        
        for row in df.itertuples(index=False):
            num_ficha = str(getattr(row, 'id', getattr(row, 'numero_ficha', ''))).split('.')[0]
            
            if not num_ficha or Ficha.query.filter_by(numero_ficha=num_ficha).first():
                continue 

            data_str = str(getattr(row, 'data_final', datetime.now().date()))
            try:
                data_final = pd.to_datetime(data_str).date()
            except:
                data_final = datetime.now().date()

            nova_ficha = Ficha(
                numero_ficha=num_ficha,
                avaliacao=converter_avaliacao(getattr(row, 'estado_geral', None)),
                autor=getattr(row, 'autor', ''),
                titulo=getattr(row, 'titulo', ''),
                registro=str(getattr(row, 'registro', '')).replace('.0', ''),
                n_chamada=str(getattr(row, 'num_chamada', '')).replace('.0', ''),
                secao_guarda=getattr(row, 'secao_guarda', ''),
                data_obra=str(getattr(row, 'data_obra', '')),
                paginas=str(getattr(row, 'num_paginas', '')).replace('.0', ''),
                dimensoes=str(getattr(row, 'dimensoes', '')),
                observacoes=getattr(row, 'observacoes', ''),
                tecnico_nome=getattr(row, 'tecnico', ''),
                data_preenchimento=data_final,
                especificacao_material={
                    'album': converter_booleano(getattr(row, 'espec_album', None)),
                    'folheto': converter_booleano(getattr(row, 'espec_folheto', None)),
                    'manuscrito': converter_booleano(getattr(row, 'espec_manuscrito', None)),
                    'planta': converter_booleano(getattr(row, 'espec_planta', None)),
                    'brochura': converter_booleano(getattr(row, 'espec_brochura', None)),
                    'gravura': converter_booleano(getattr(row, 'espec_gravura', None)),
                    'mapa': converter_booleano(getattr(row, 'espec_mapa', None)),
                    'pergaminho': converter_booleano(getattr(row, 'espec_pergaminho_scroll', None)),
                    'certificado': converter_booleano(getattr(row, 'espec_certificado', None)),
                    'impresso': converter_booleano(getattr(row, 'espec_impresso', None)),
                    'partitura': converter_booleano(getattr(row, 'espec_partitura', None)),
                    'desenho': converter_booleano(getattr(row, 'espec_desenho', None)),
                    'livro': converter_booleano(getattr(row, 'espec_livro', None)),
                    'periodico': converter_booleano(getattr(row, 'espec_periodico', None)),
                    'outro_texto': ''
                },
                tipo_suporte={
                    'couche': converter_booleano(getattr(row, 'sup_papel_couche', None)),
                    'jornal': converter_booleano(getattr(row, 'sup_papel_jornal', None)),
                    'feito_mao': converter_booleano(getattr(row, 'sup_papel_feito_a_mao', None)),
                    'madeira': converter_booleano(getattr(row, 'sup_papel_madeira', None)),
                    'trapo': converter_booleano(getattr(row, 'sup_papel_trapo', None)),
                    'marmorizado': converter_booleano(getattr(row, 'sup_papel_marmorizado', None))
                },
                estado_conservacao={
                    'encadernada': converter_booleano(getattr(row, 'enc_tipo', None) == 'Encadernada' or not converter_booleano(getattr(row, 'sem_encadernacao', None))),
                    'sem_encadernacao': converter_booleano(getattr(row, 'sem_encadernacao', None)),
                    'tapa_madeira': converter_booleano(getattr(row, 'tapa_madeira', None)),
                    'tapa_papelao': converter_booleano(getattr(row, 'tapa_papelao', None)),
                    'capa_couro': converter_booleano(getattr(row, 'capa_couro', None)),
                    'capa_tecido': converter_booleano(getattr(row, 'capa_tecido', None))
                },
                deterioracoes={
                    'abrasao': converter_booleano(getattr(row, 'det_enc_abrasao', None)),
                    'arranhao': converter_booleano(getattr(row, 'det_enc_arranhao', None)),
                    'costura_fragil': converter_booleano(getattr(row, 'det_enc_costura_fragilizada', None)),
                    'descoloracao': converter_booleano(getattr(row, 'det_enc_descoloracao', None)),
                    'lombada_quebrada': converter_booleano(getattr(row, 'det_enc_lombada_quebrada', None)),
                    'mancha': converter_booleano(getattr(row, 'det_enc_mancha', None)) or converter_booleano(getattr(row, 'det_miolo_mancha', None)),
                    'rompimento': converter_booleano(getattr(row, 'det_enc_rompimento', None)),
                    'sujidades': converter_booleano(getattr(row, 'det_enc_sujidades', None)) or converter_booleano(getattr(row, 'det_miolo_sujidade', None)),
                    'fungos': converter_booleano(getattr(row, 'det_miolo_fungos', None)),
                    'oxidacao': converter_booleano(getattr(row, 'det_miolo_oxidacao', None))
                },
                tratamento_planos={
                    'diagnostico': converter_booleano(getattr(row, 'trat_plano_diagnostico', None)),
                    'higienizacao': converter_booleano(getattr(row, 'trat_plano_higienizacao', None)),
                    'retirada_sujidades': converter_booleano(getattr(row, 'trat_plano_retirada_de_sujidades_extrinsecas', None)),
                    'retirada_fitas': converter_booleano(getattr(row, 'trat_plano_retirada_de_fitas_adesivas', None)),
                    'desacidificacao': converter_booleano(getattr(row, 'trat_plano_desacidificacao_a_seco', None)),
                    'arrefecimento': converter_booleano(getattr(row, 'trat_plano_arrefecimento_de_manchas', None)),
                    'reestruturacao': converter_booleano(getattr(row, 'trat_plano_reestruturacao', None)),
                    'remendos': converter_booleano(getattr(row, 'trat_plano_remendos', None)),
                    'enxertos': converter_booleano(getattr(row, 'trat_plano_enxertos', None)),
                    'velaturas': converter_booleano(getattr(row, 'trat_plano_velaturas', None)),
                    'planificacao': converter_booleano(getattr(row, 'trat_plano_planificacao', None)),
                    'acondicionamento': converter_booleano(getattr(row, 'trat_plano_acondicionamento', None)),
                    'portfolio': converter_booleano(getattr(row, 'trat_plano_portfolio', None)),
                    'envelope': converter_booleano(getattr(row, 'trat_plano_envelope', None)),
                    'passe_partout': converter_booleano(getattr(row, 'trat_plano_passe_partout', None)),
                    'pasta': converter_booleano(getattr(row, 'trat_plano_pasta', None)),
                    'jaqueta': converter_booleano(getattr(row, 'trat_plano_jaqueta_de_poliester', None))
                },
                tratamento_volumes={
                    'fumigacao': converter_booleano(getattr(row, 'trat_vol_fumigacao', None)),
                    'higienizacao': converter_booleano(getattr(row, 'trat_vol_higienizacao', None)),
                    'reestruturacao': converter_booleano(getattr(row, 'trat_vol_reestruturacao', None)),
                    'lombada': converter_booleano(getattr(row, 'trat_vol_lombada', None)),
                    'lombada_capa': converter_booleano(getattr(row, 'trat_vol_lombada_e_capa', None))
                }
            )

            db.session.add(nova_ficha)
            imported_count += 1
        
        db.session.commit()
        flash(f'Sucesso! {imported_count} fichas importadas.', 'success')
        return redirect(url_for('listar_acervo'))

    except Exception as e:
        db.session.rollback()
        print(f"ERRO CRÍTICO NA IMPORTAÇÃO: {e}")
        flash(f'Erro ao processar: {str(e)}', 'danger')
        return redirect(url_for('listar_acervo'))

@app.route('/exportar')
def exportar_planilha():
    try:
        fichas = Ficha.query.all()
        
        if not fichas:
            flash('Não há fichas para exportar.', 'warning')
            return redirect(url_for('listar_acervo'))

        lista_dados = []

        def processar_multiplos(dados_json, mapa_nomes):
            if not dados_json: return ""
            itens_encontrados = []
            
            for chave_db, nome_legivel in mapa_nomes.items():
                if dados_json.get(chave_db):
                    itens_encontrados.append(nome_legivel)
            
            outro = dados_json.get('outro_texto')
            if outro and str(outro).strip():
                itens_encontrados.append(f"Outro: {outro}")
            
            return ", ".join(itens_encontrados)

        def traduzir_avaliacao(valor):
            if valor == 1: return "Bom"
            if valor == 3: return "Mau"
            return "Regular"

        mapa_material = {
            'album': 'Álbum', 'folheto': 'Folheto', 'manuscrito': 'Manuscrito',
            'planta': 'Planta', 'brochura': 'Brochura', 'gravura': 'Gravura',
            'mapa': 'Mapa', 'pergaminho': 'Pergaminho', 'certificado': 'Certificado',
            'impresso': 'Impresso', 'partitura': 'Partitura', 'desenho': 'Desenho',
            'livro': 'Livro', 'periodico': 'Periódico'
        }

        mapa_suporte = {
            'couche': 'Papel Couchê', 'jornal': 'Papel Jornal',
            'feito_mao': 'Papel Feito à Mão', 'madeira': 'Papel Madeira',
            'trapo': 'Papel de Trapo', 'marmorizado': 'Papel Marmorizado'
        }

        mapa_estado = {
            'encadernada': 'Encadernada', 'sem_encadernacao': 'Sem Encadernação',
            'inteira': 'Enc. Inteira', 'meia_com_cantos': '½ com cantos',
            'meia_sem_cantos': '½ sem cantos', 'capa_couro': 'Capa Couro',
            'capa_tecido': 'Capa Tecido', 'tapa_madeira': 'Tapa Madeira',
            'tapa_papelao': 'Tapa Papelão'
        }

        mapa_deterioracoes = {
            'abrasao': 'Abrasão', 'costura_fragil': 'Costura Fragilizada',
            'mancha': 'Mancha', 'rompimento': 'Rompimento', 'arranhao': 'Arranhão',
            'descoloracao': 'Descoloração', 'perda_lombada': 'Perda de Lombada',
            'sujidades': 'Sujidades', 'fungos': 'Fungos', 
            'oxidacao': 'Oxidação', 'lombada_quebrada': 'Lombada Quebrada'
        }

        mapa_plano = {
            'diagnostico': 'Diagnóstico', 'higienizacao': 'Higienização',
            'retirada_sujidades': 'Retirada Sujidades', 'retirada_fitas': 'Retirada Fitas',
            'desacidificacao': 'Desacidificação', 'arrefecimento': 'Arrefecimento',
            'reestruturacao': 'Reestruturação', 'remendos': 'Remendos',
            'enxertos': 'Enxertos', 'velaturas': 'Velaturas',
            'planificacao': 'Planificação', 'acondicionamento': 'Acondicionamento',
            'portfolio': 'Portfólio', 'passe_partout': 'Passe-partout',
            'pasta': 'Pasta', 'envelope': 'Envelope', 'jaqueta': 'Jaqueta'
        }

        mapa_volume = {
            'fumigacao': 'Fumigação', 'fungos': 'Trat. Fungos', 'insetos': 'Trat. Insetos',
            'higienizacao': 'Higienização', 'trincha': 'Trincha',
            'reestruturacao': 'Reestruturação', 'lombada': 'Lombada',
            'lombada_capa': 'Lombada e Capa', 'folhas': 'Folhas',
            'encadernacao': 'Encadernação', 'inteira': 'Inteira',
            'meia_sem_cantos': '½ Sem cantos', 'costura': 'Costura',
            'douracao': 'Douração', 'punho': 'A Punho', 'maquina': 'À Máquina',
            'acondicionamento': 'Acondicionamento', 'caixa_cruz': 'Caixa Cruz',
            'caixa_cadarco': 'Caixa Cadarço'
        }

        for f in fichas:
            linha = {
                'ID': f.numero_ficha,
                'Avaliação': traduzir_avaliacao(f.avaliacao),
                'Autor': f.autor,
                'Título': f.titulo,
                'Registro': f.registro,
                'Nº Chamada': f.n_chamada,
                'Seção': f.secao_guarda,
                'Data Obra': f.data_obra,
                'Páginas': f.paginas,
                'Dimensões': f.dimensoes,
                'Especificação do Material': processar_multiplos(f.especificacao_material, mapa_material),
                'Tipo de Suporte': processar_multiplos(f.tipo_suporte, mapa_suporte),
                'Estado de Conservação': processar_multiplos(f.estado_conservacao, mapa_estado),
                'Deteriorações': processar_multiplos(f.deterioracoes, mapa_deterioracoes),
                'Tratamento (Planos)': processar_multiplos(f.tratamento_planos, mapa_plano),
                'Tratamento (Volumes)': processar_multiplos(f.tratamento_volumes, mapa_volume),
                'Observações': f.observacoes,
                'Técnico': f.tecnico_nome,
                'Data Preenchimento': f.data_preenchimento,
                'Foto': f.foto_path
            }
            lista_dados.append(linha)

        df = pd.DataFrame(lista_dados)
        
        colunas_ordem = ['ID', 'Título', 'Autor', 'Avaliação', 'Especificação do Material', 
                         'Tipo de Suporte', 'Estado de Conservação', 'Deteriorações', 
                         'Tratamento (Planos)', 'Tratamento (Volumes)', 'Observações', 'Técnico']
        cols_existentes = [c for c in colunas_ordem if c in df.columns]
        cols_restantes = [c for c in df.columns if c not in cols_existentes]
        df = df[cols_existentes + cols_restantes]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Acervo Consolidado')
            
            worksheet = writer.sheets['Acervo Consolidado']
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                if length > 50: length = 50
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
        
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Relatorio_Resumido.xlsx'
        )

    except Exception as e:
        print(f"Erro na exportação: {e}")
        flash(f'Erro ao gerar planilha: {str(e)}', 'danger')
        return redirect(url_for('listar_acervo'))