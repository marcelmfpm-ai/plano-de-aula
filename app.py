from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.shared import Pt
import io
import os
import json
import firebase_admin
from firebase_admin import credentials, firestore
from datetime import datetime

BASE_DIR_EARLY = os.path.dirname(os.path.abspath(__file__))

_firebase_creds_env = os.environ.get('FIREBASE_CREDENTIALS')
if _firebase_creds_env:
    _cred = credentials.Certificate(json.loads(_firebase_creds_env))
else:
    _cred_path = os.path.join(BASE_DIR_EARLY, 'plano-de-aula-4a383-firebase-adminsdk-fbsvc-bd042c3e0f.json')
    _cred = credentials.Certificate(_cred_path)

firebase_admin.initialize_app(_cred)
db = firestore.client()

VERDANA = 'Verdana'
TAMANHO = Pt(12)


def _fmt(run, bold=False):
    run.font.name = VERDANA
    run.font.size = TAMANHO
    if bold:
        run.bold = True


def _set_spacing_15(para):
    pPr = para._p.get_or_add_pPr()
    # espaçamento 1,5
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:line'), '360')
    spacing.set(qn('w:lineRule'), 'auto')
    # justificado via XML (mais confiável que a propriedade Python)
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), 'both')


def inserir_tabela_apos(doc, paragrafo_ref, rows):
    tbl = doc.add_table(rows=1 + len(rows), cols=3)
    tbl.style = 'Table Grid'

    cabecalhos = ['Etapa', 'Tempo Estimado', 'Tempo da Aula']
    for j, cab in enumerate(cabecalhos):
        cell = tbl.rows[0].cells[j]
        cell.text = cab
        run = cell.paragraphs[0].runs[0]
        _fmt(run, bold=True)

    for i, (etapa, tempo, acum) in enumerate(rows):
        linha = tbl.rows[i + 1]
        for cell, txt in zip(linha.cells, [etapa, tempo, acum]):
            cell.text = txt
            for run in cell.paragraphs[0].runs:
                _fmt(run)

    paragrafo_ref._element.addnext(tbl._tbl)
    return tbl


def inserir_linha_apos(paragrafo_ref, texto, estilo='Normal'):
    novo_xml = OxmlElement('w:p')
    paragrafo_ref._element.addnext(novo_xml)
    p = DocxParagraph(novo_xml, paragrafo_ref._parent)
    try:
        p.style = estilo
    except Exception:
        pass
    _set_spacing_15(p)
    if texto:
        _fmt(p.add_run(texto))
    return p


def _paragrafo_formatado(ultimo, paragrafo_ref):
    """Cria parágrafo após 'ultimo', com pai de 'paragrafo_ref', com 1,5 e justificado."""
    novo_xml = OxmlElement('w:p')
    ultimo._element.addnext(novo_xml)
    p = DocxParagraph(novo_xml, paragrafo_ref._parent)
    try:
        p.style = 'Normal'
    except Exception:
        pass
    _set_spacing_15(p)
    return p


def inserir_lista_apos(paragrafo_ref, itens, ret=False):
    ultimo = paragrafo_ref
    for item in itens:
        if item.strip():
            novo_xml = OxmlElement('w:p')
            ultimo._element.addnext(novo_xml)
            p = DocxParagraph(novo_xml, paragrafo_ref._parent)
            try:
                p.style = 'Normal'
            except Exception:
                pass
            _set_spacing_15(p)
            _fmt(p.add_run(item.strip()))
            ultimo = p
    if ret:
        return ultimo


def inserir_lista_formatada(paragrafo_ref, itens, ret=False):
    """Objetivos/ênfase: ● negrito + primeira palavra negrito, 1,5 e justificado."""
    ultimo = paragrafo_ref
    for item in itens:
        texto = item.strip()
        if not texto:
            continue
        p = _paragrafo_formatado(ultimo, paragrafo_ref)
        _fmt(p.add_run('● '), bold=True)
        partes = texto.split(' ', 1)
        _fmt(p.add_run(partes[0]), bold=True)
        if len(partes) > 1:
            _fmt(p.add_run(' ' + partes[1]))
        ultimo = p
    if ret:
        return ultimo


def inserir_lista_diretrizes(paragrafo_ref, itens, ret=False):
    """Diretrizes: ● negrito + texto normal, 1,5 e justificado."""
    ultimo = paragrafo_ref
    for item in itens:
        texto = item.strip()
        if not texto:
            continue
        p = _paragrafo_formatado(ultimo, paragrafo_ref)
        _fmt(p.add_run('● '), bold=True)
        _fmt(p.add_run(texto))
        ultimo = p
    if ret:
        return ultimo

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'ESTRUTURA DO PLANO DE AULA.docx')


def _remover_imagens_apos_material(doc):
    parags = list(doc.paragraphs)
    inicio = None
    fim = None
    for i, p in enumerate(parags):
        t = p.text.strip()
        if inicio is None and '3.2' in t:
            inicio = i
        elif inicio is not None and t.startswith('4.'):
            fim = i
            break
    if inicio is None:
        return
    fim = fim if fim is not None else len(parags)
    para_remover = []
    for p in parags[inicio + 1:fim]:
        if (p._p.find('.//' + qn('w:drawing')) is not None or
                p._p.find('.//' + qn('w:pict')) is not None):
            para_remover.append(p._p)
    for p_xml in para_remover:
        parent = p_xml.getparent()
        if parent is not None:
            parent.remove(p_xml)


def substituir_paragrafo(para, substituicoes):
    texto_original = para.text
    if not any(k in texto_original for k in substituicoes):
        return

    novo_texto = texto_original
    for antigo, novo in substituicoes.items():
        novo_texto = novo_texto.replace(antigo, novo)

    if para.runs:
        para.runs[0].text = novo_texto
        for run in para.runs[1:]:
            run.text = ''


@app.route('/')
def index():
    tmpl = app.jinja_env.get_template('index.html')
    print('TEMPLATE PATH:', tmpl.filename)
    return render_template('index.html')


@app.route('/gerar', methods=['POST'])
def gerar():
    data = request.form
    doc = Document(TEMPLATE_PATH)
    _remover_imagens_apos_material(doc)

    substituicoes = {
        '(CARGO)':             data.get('cargo', '').upper(),
        '(MODULO)':            data.get('modulo', ''),
        '(UNIDADE)':           data.get('unidade', ''),
        '(HORAS AULA)':        data.get('horas_aula', ''),
        '(CARGA HORÁRIA)':     data.get('carga_horaria', '').replace(' min', ''),
        '(MODULO DESCRIÇÃO)':  data.get('modulo_descricao', ''),
        '(ATIVIDADE DESCRIÇÃO)': data.get('atividade_descricao', ''),
        '(DESCRIÇÃO DA AULA)':   data.get('descricao_aula', ''),
        '(LOCAL DA AULA)':       data.get('local_aula', ''),
        '(DIVISÃO DA TURMA)':    data.get('divisao_turma_outro', '').strip()
                                 if data.get('divisao_turma') == 'Outro'
                                 else data.get('divisao_turma', ''),
        '(GRUPO DA TURMA)':      data.get('grupo_turma', ''),
    }

    humano_itens = []

    # Professores APC — agrupa subtipos e conta cada um
    from collections import Counter
    tipos_apc = data.getlist('prof_apc_tipo')
    contagem_apc = Counter(tipos_apc)
    for tipo, qtd in contagem_apc.items():
        humano_itens.append(f'{qtd:02d} Professor(es) APC — {tipo}')

    # Demais professores e monitores
    campos_humano = [
        ('prof_damaz',    'Professor(es) DAMAZ'),
        ('prof_dciber',   'Professor(es) DCIBER'),
        ('prof_epol',     'Professor(es) Epol'),
        ('monitor_apc',   'Monitor(es) de APC'),
        ('monitor_sala',  'Monitor(es) Sala'),
        ('monitor_epol',  'Monitor(es) Epol'),
        ('monitor_damaz', 'Monitor(es) Damaz'),
        ('monitor_dciber','Monitor(es) DCIBER'),
    ]
    for campo, label in campos_humano:
        qtd = int(data.get(campo, 0) or 0)
        if qtd > 0:
            humano_itens.append(f'{qtd:02d} {label}')

    tipos_aux   = data.getlist('auxiliar_tipo')
    outros_aux  = data.getlist('auxiliar_outro')
    contagem_aux = Counter()
    outros_livres = []
    for i, tipo in enumerate(tipos_aux):
        if tipo == 'Outro':
            txt = outros_aux[i] if i < len(outros_aux) else ''
            if txt.strip():
                outros_livres.append(txt.strip())
        else:
            contagem_aux[tipo] += 1
    for tipo, qtd in contagem_aux.items():
        humano_itens.append(f'{qtd:02d} Auxiliar(es) — {tipo}')
    for txt in outros_livres:
        humano_itens.append(f'Auxiliar — {txt}')

    mat_qtds   = data.getlist('mat_qtd')
    mat_labels = data.getlist('mat_label')
    material_itens = []
    for qtd, label in zip(mat_qtds, mat_labels):
        n = int(qtd or 0)
        if n > 0:
            material_itens.append(f'{n:02d} {label}')

    objetivo_itens = [f.strip() for f in data.getlist('obj_verbo') if f.strip()]
    enfase_itens   = [f.strip() for f in data.getlist('enfase_noun') if f.strip()]

    etapa_nomes    = data.getlist('etapa_nome')
    etapa_outros   = data.getlist('etapa_nome_outro')
    etapa_tempos   = data.getlist('etapa_tempo')
    etapa_detalhes = data.getlist('etapa_detalhe')
    estrutura_rows    = []
    estrutura_detalhes = []
    acumulado = 0
    for i, nome in enumerate(etapa_nomes):
        nome_final = etapa_outros[i].strip() if nome == '__outro__' and i < len(etapa_outros) else nome
        tempo = int(etapa_tempos[i] or 0) if i < len(etapa_tempos) else 0
        acumulado += tempo
        detalhe = etapa_detalhes[i].strip() if i < len(etapa_detalhes) else ''
        if nome_final:
            estrutura_rows.append((nome_final, f'{tempo} min', f'{acumulado} min'))
            estrutura_detalhes.append((nome_final, detalhe))

    verificacao_itens  = [v.strip() for v in data.getlist('verificacao')       if v.strip()]
    ressalva_itens     = [v.strip() for v in data.getlist('ressalva')          if v.strip()]
    estrategia_intro   = data.get('estrategia_intro', '').strip()
    estrategia_bullets = [v.strip() for v in data.getlist('estrategia_bullet') if v.strip()]

    for para in doc.paragraphs:
        substituir_paragrafo(para, substituicoes)

    for para in doc.paragraphs:
        texto = para.text
        if 'OBJETIVOS DE APRENDIZAGEM' in texto:
            ref = inserir_linha_apos(para, 'Ao final da aula, os alunos deverão ser capazes de:', 'Normal')
            ref = inserir_lista_formatada(ref, objetivo_itens, ret=True)
            if enfase_itens:
                ref = inserir_linha_apos(ref, '', 'Normal')
                enfase_titulo = inserir_linha_apos(ref, 'Ênfase:', 'Normal')
                inserir_lista_formatada(enfase_titulo, enfase_itens)
        elif '3.1 HUMANO' in texto:
            inserir_lista_apos(para, humano_itens)
        elif '3.2 MATERIAL' in texto:
            inserir_lista_apos(para, material_itens)
        elif 'ESTRATÉGIA DE ENSINO' in texto:
            ref = para
            if estrategia_intro:
                ref = inserir_linha_apos(ref, estrategia_intro, 'Normal')
            inserir_lista_diretrizes(ref, estrategia_bullets)
        elif 'ESTRUTURA GERAL DA AULA' in texto:
            if estrutura_rows:
                tbl = inserir_tabela_apos(doc, para, estrutura_rows)
                ultimo = tbl._tbl
                for idx, (nome_etapa, detalhe) in enumerate(estrutura_detalhes):
                    p_blank = OxmlElement('w:p')
                    ultimo.addnext(p_blank)
                    p_tit_xml = OxmlElement('w:p')
                    p_blank.addnext(p_tit_xml)
                    p_tit = DocxParagraph(p_tit_xml, para._parent)
                    _fmt(p_tit.add_run(f'4.{idx + 1} {nome_etapa}'), bold=True)
                    p_det_xml = OxmlElement('w:p')
                    p_tit_xml.addnext(p_det_xml)
                    p_det = DocxParagraph(p_det_xml, para._parent)
                    if detalhe:
                        _fmt(p_det.add_run(detalhe))
                    ultimo = p_det_xml
        elif 'Ressalvas Did' in texto:
            inserir_lista_apos(para, ressalva_itens)
        elif 'VERIFICAÇÃO DE APRENDIZAGEM' in texto:
            ref = para
            for item in verificacao_itens:
                ref = inserir_linha_apos(ref, item, 'Normal')

    # Aplica espaçamento 1,5 e justificado em todos os parágrafos do documento
    for para in doc.paragraphs:
        _set_spacing_15(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _set_spacing_15(para)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    mod = data.get('modulo', 'XX').zfill(2)
    uni = data.get('unidade', 'XX').zfill(2)
    filename = f"Plano_Aula_M{mod}_U{uni}.docx"

    return send_file(
        buf,
        download_name=filename,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@app.route('/salvar', methods=['POST'])
def salvar():
    data = request.json
    nome = data.get('nome', '').strip() or 'Sem título'
    doc_ref = db.collection('projetos').document()
    doc_ref.set({
        'nome': nome,
        'usuario': data.get('usuario', ''),
        'dados': data.get('dados', {}),
        'status': data.get('status', 'rascunho'),
        'criado_em': datetime.now().isoformat(),
        'atualizado_em': datetime.now().isoformat(),
    })
    return jsonify({'id': doc_ref.id, 'nome': nome})


@app.route('/projetos', methods=['GET'])
def listar_projetos():
    docs = db.collection('projetos').order_by('atualizado_em', direction=firestore.Query.DESCENDING).stream()
    result = []
    for doc in docs:
        d = doc.to_dict()
        result.append({
            'id': doc.id,
            'nome': d.get('nome', ''),
            'usuario': d.get('usuario', ''),
            'status': d.get('status', ''),
            'atualizado_em': d.get('atualizado_em', ''),
        })
    return jsonify(result)


@app.route('/projeto/<doc_id>', methods=['GET'])
def carregar_projeto(doc_id):
    doc = db.collection('projetos').document(doc_id).get()
    if not doc.exists:
        return jsonify({'erro': 'Projeto não encontrado'}), 404
    d = doc.to_dict()
    return jsonify({'id': doc.id, 'nome': d.get('nome'), 'dados': d.get('dados', {}), 'status': d.get('status')})


@app.route('/projeto/<doc_id>', methods=['PUT'])
def atualizar_projeto(doc_id):
    data = request.json
    db.collection('projetos').document(doc_id).update({
        'nome': data.get('nome', ''),
        'dados': data.get('dados', {}),
        'status': data.get('status', 'rascunho'),
        'atualizado_em': datetime.now().isoformat(),
    })
    return jsonify({'ok': True})


@app.route('/projeto/<doc_id>', methods=['DELETE'])
def deletar_projeto(doc_id):
    db.collection('projetos').document(doc_id).delete()
    return jsonify({'ok': True})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
