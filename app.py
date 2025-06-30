# ***************************************** version ***********************************************

# *************************************************************************************************

# Bibliotecas
import io
import os
import re
from datetime import datetime
from time import time

import pandas as pd
from flask import (
    Flask,
    jsonify,
    render_template,
    request,
    send_file,
    send_from_directory,
)
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurações
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'csv', 'xlsx'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# Colunas obrigatórias
REQUIRED_COLUMNS = [
    'IDENTIFICADOR_LOCAL',
    'DOCUMENTO_PACIENTE',
    'DATA_SOLICITACAO',
    'CNES_SOLICITANTE',
    'CNES_REGULADOR',
    'CODIGO_SIGTAP',
    'CBO',
    'CID10',
    'CODIGO_MODALIDADE_ASSISTENCIAL',
    'CODIGO_CARTER_SOLICITACAO',
    'STATUS',
    'DATA_AUTORIZACAO',
    'DATA_EXECUCAO',
    'CNES_EXECUTANTE',
]

# Definição dos agrupamentos (seu dicionário completo de agrupamentos aqui)
agrupamentos = {
    '0901010014': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA INICIAL DE CÂNCER DE MAMA',
        'itens_obrigatorios': [
            {'codigo': '0204030030', 'descricao': 'MAMOGRAFIA'},
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0205020097',
                'descricao': 'ULTRASSONOGRAFIA MAMARIA BILATERAL',
            }
        ],
    },
    '0901010090': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA DE CÂNCER DE MAMA - I',
        'itens_obrigatorios': [
            {
                'codigo': '0203010043',
                'descricao': 'EXAME CITOPATOLOGICO DE MAMA',
            },
            {
                'codigo': '0201010585',
                'descricao': 'PUNÇÃO ASPIRATIVA DE MAMA POR AGULHA FINA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0201010569',
                'descricao': 'BIOPSIA/EXERESE DE NÓDULO DE MAMA',
            }
        ],
    },
    '0901010103': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA DE CÂNCER DE MAMA-II',
        'itens_obrigatorios': [
            {
                'codigo': '0201010607',
                'descricao': 'PUNÇÃO DE MAMA POR AGULHA GROSSA',
            },
            {
                'codigo': '0203020065',
                'descricao': 'EXAME ANATOMOPATOLOGICO DE MAMA - BIOPSIA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0201010569',
                'descricao': 'BIOPSIA/EXERESE DE NÓDULO DE MAMA',
            }
        ],
    },
    '0901010057': {
        'nome': 'OCI INVESTIGAÇÃO DIAGNÓSTICA DE CÂNCER DE COLO DE ÚTERO',
        'itens_obrigatorios': [
            {'codigo': '0201010666', 'descricao': 'BIOPSIA DO COLO UTERINO'},
            {
                'codigo': '0203020081',
                'descricao': 'EXAME ANATOMO-PATOLOGICO DO COLO UTERINO - BIOPSIA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211040029', 'descricao': 'COLPOSCOPIA'}
        ],
    },
    '0901010111': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA E TERAPÊUTICA DE CÂNCER DE COLO DO ÚTERO-I',
        'itens_obrigatorios': [
            {
                'codigo': '0203020022',
                'descricao': 'EXAME ANATOMO-PATOLOGICO DO COLO UTERINO - PECA CIRURGICA',
            },
            {
                'codigo': '0409060089',
                'descricao': 'EXCISÃO TIPO I DO COLO UTERINO',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211040029', 'descricao': 'COLPOSCOPIA'}
        ],
    },
    '0901010120': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA E TERAPÊUTICA DE CÂNCER DE COLO DO ÚTERO-II',
        'itens_obrigatorios': [
            {
                'codigo': '0203020022',
                'descricao': 'EXAME ANATOMO-PATOLOGICO DO COLO UTERINO - PECA CIRURGICA',
            },
            {
                'codigo': '0409060305',
                'descricao': 'EXCISÃO TIPO 2 DO COLO UTERINO',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211040029', 'descricao': 'COLPOSCOPIA'}
        ],
    },
    '0901010049': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA DE CÂNCER DE PRÓSTATA',
        'itens_obrigatorios': [
            {'codigo': '0201010410', 'descricao': 'BIÓPSIA DE PRÓSTATA'},
            {
                'codigo': '0203020030',
                'descricao': 'EXAME ANATOMO-PATOLÓGICO PARA CONGELAMENTO / PARAFINA POR PEÇA CIRURGICA OU POR BIOPSIA (EXCETO COLO UTERINO E MAMA)',
            },
            {
                'codigo': '0205020119',
                'descricao': 'ULTRASSONOGRAFIA DE PROSTATA (VIA TRANSRETAL)',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0901010073': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA DE CÂNCER GÁSTRICO',
        'itens_obrigatorios': [
            {
                'codigo': '0209010037',
                'descricao': 'ESOFAGOGASTRODUODENOSCOPIA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0203020030',
                'descricao': 'EXAME ANATOMO-PATOLÓGICO PARA CONGELAMENTO / PARAFINA POR PEÇA CIRURGICA OU POR BIOPSIA (EXCETO COLO UTERINO E MAMA)',
            }
        ],
    },
    '0901010081': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA DE CÂNCER COLORRETAL',
        'itens_obrigatorios': [
            {'codigo': '0209010029', 'descricao': 'COLONOSCOPIA (COLOSCOPIA)'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0203020030',
                'descricao': 'EXAME ANATOMO-PATOLÓGICO PARA CONGELAMENTO / PARAFINA POR PEÇA CIRURGICA OU POR BIOPSIA (EXCETO COLO UTERINO E MAMA)',
            }
        ],
    },
    '0902010018': {
        'nome': 'OCI AVALIAÇÃO DE RISCO CIRÚRGICO',
        'itens_obrigatorios': [
            {'codigo': '0211020036', 'descricao': 'ELETROCARDIOGRAMA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0202010279', 'descricao': 'DOSAGEM DE COLESTEROL HDL'},
            {'codigo': '0202010287', 'descricao': 'DOSAGEM DE COLESTEROL LDL'},
            {
                'codigo': '0202010295',
                'descricao': 'DOSAGEM DE COLESTEROL TOTAL',
            },
            {'codigo': '0202010317', 'descricao': 'DOSAGEM DE CREATININA'},
            {'codigo': '0202010473', 'descricao': 'DOSAGEM DE GLICOSE'},
            {
                'codigo': '0202010503',
                'descricao': 'DOSAGEM DE HEMOGLOBINA GLICOSILADA',
            },
            {'codigo': '0202010600', 'descricao': 'DOSAGEM DE POTASSIO'},
            {'codigo': '0202010635', 'descricao': 'DOSAGEM DE SODIO'},
            {
                'codigo': '0202010643',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-OXALACETICA (TGO)',
            },
            {
                'codigo': '0202010651',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-PIRUVICA (TGP)',
            },
            {'codigo': '0202010678', 'descricao': 'DOSAGEM DE TRIGLICERIDEOS'},
            {'codigo': '0202010694', 'descricao': 'DOSAGEM DE UREIA'},
            {'codigo': '0202020380', 'descricao': 'HEMOGRAMA COMPLETO'},
            {
                'codigo': '0204030153',
                'descricao': 'RADIOGRAFIA DE TORAX (PA E PERFIL)',
            },
        ],
    },
    '0902010026': {
        'nome': 'OCI AVALIAÇÃO CARDIOLÓGICA',
        'itens_obrigatorios': [
            {
                'codigo': '0204030153',
                'descricao': 'RADIOGRAFIA DE TORAX (PA E PERFIL)',
            },
            {'codigo': '0211020036', 'descricao': 'ELETROCARDIOGRAMA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0202010279', 'descricao': 'DOSAGEM DE COLESTEROL HDL'},
            {'codigo': '0202010287', 'descricao': 'DOSAGEM DE COLESTEROL LDL'},
            {
                'codigo': '0202010295',
                'descricao': 'DOSAGEM DE COLESTEROL TOTAL',
            },
            {'codigo': '0202010317', 'descricao': 'DOSAGEM DE CREATININA'},
            {'codigo': '0202010473', 'descricao': 'DOSAGEM DE GLICOSE'},
            {
                'codigo': '0202010503',
                'descricao': 'DOSAGEM DE HEMOGLOBINA GLICOSILADA',
            },
            {'codigo': '0202010600', 'descricao': 'DOSAGEM DE POTASSIO'},
            {'codigo': '0202010635', 'descricao': 'DOSAGEM DE SODIO'},
            {
                'codigo': '0202010643',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-OXALACETICA (TGO)',
            },
            {
                'codigo': '0202010651',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-PIRUVICA (TGP)',
            },
            {'codigo': '0202010678', 'descricao': 'DOSAGEM DE TRIGLICERIDEOS'},
            {'codigo': '0202010694', 'descricao': 'DOSAGEM DE UREIA'},
            {'codigo': '0202020380', 'descricao': 'HEMOGRAMA COMPLETO'},
            {
                'codigo': '0205010032',
                'descricao': 'ECOCARDIOGRAFIA TRANSTORACICA',
            },
        ],
    },
    '0902010034': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA INICIAL - SÍNDROME CORONARIANA CRÔNICA',
        'itens_obrigatorios': [
            {'codigo': '0211020036', 'descricao': 'ELETROCARDIOGRAMA'},
            {
                'codigo': '0211020060',
                'descricao': 'TESTE DE ESFORCO / TESTE ERGOMETRICO',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0202010279', 'descricao': 'DOSAGEM DE COLESTEROL HDL'},
            {'codigo': '0202010287', 'descricao': 'DOSAGEM DE COLESTEROL LDL'},
            {
                'codigo': '0202010295',
                'descricao': 'DOSAGEM DE COLESTEROL TOTAL',
            },
            {'codigo': '0202010317', 'descricao': 'DOSAGEM DE CREATININA'},
            {'codigo': '0202010473', 'descricao': 'DOSAGEM DE GLICOSE'},
            {
                'codigo': '0202010503',
                'descricao': 'DOSAGEM DE HEMOGLOBINA GLICOSILADA',
            },
            {'codigo': '0202010600', 'descricao': 'DOSAGEM DE POTASSIO'},
            {'codigo': '0202010635', 'descricao': 'DOSAGEM DE SODIO'},
            {
                'codigo': '0202010643',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-OXALACETICA (TGO)',
            },
            {
                'codigo': '0202010651',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-PIRUVICA (TGP)',
            },
            {'codigo': '0202010678', 'descricao': 'DOSAGEM DE TRIGLICERIDEOS'},
            {'codigo': '0202010694', 'descricao': 'DOSAGEM DE UREIA'},
            {'codigo': '0202020380', 'descricao': 'HEMOGRAMA COMPLETO'},
            {
                'codigo': '0205010032',
                'descricao': 'ECOCARDIOGRAFIA TRANSTORACICA',
            },
        ],
    },
    '0902010042': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA I – SÍNDROME CORONARIANA CRÔNICA',
        'itens_obrigatorios': [
            {
                'codigo': '0205010016',
                'descricao': 'ECOCARDIOGRAFIA DE ESTRESSE',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0902010050': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA II – SÍNDROME CORONARIANA CRÔNICA',
        'itens_obrigatorios': [
            {
                'codigo': '0208010025',
                'descricao': 'CINTILOGRAFIA DE MIOCARDIO P/ AVALIACAO DA PERFUSAO EM SITUACAO DE ESTRESSE (MINIMO 3 PROJECOES)',
            },
            {
                'codigo': '0208010033',
                'descricao': 'CINTILOGRAFIA DE MIOCARDIO P/ AVALIACAO DA PERFUSAO EM SITUACAO DE REPOUSO (MINIMO 3 PROJECOES)',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0902010069': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA - INSUFICIÊNCIA CARDÍACA',
        'itens_obrigatorios': [
            {
                'codigo': '0202010791',
                'descricao': 'DOSAGEM DE PEPTÍDEOS NATRIURÉTICOS TIPO B (BNP E NT-PROBNP)',
            },
            {'codigo': '0211020036', 'descricao': 'ELETROCARDIOGRAMA'},
            {
                'codigo': '0211020044',
                'descricao': 'MONITORAMENTO PELO SISTEMA HOLTER 24 HS (3 CANAIS)',
            },
            {
                'codigo': '0211020060',
                'descricao': 'TESTE DE ESFORCO / TESTE ERGOMETRICO',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0202010279', 'descricao': 'DOSAGEM DE COLESTEROL HDL'},
            {'codigo': '0202010287', 'descricao': 'DOSAGEM DE COLESTEROL LDL'},
            {
                'codigo': '0202010295',
                'descricao': 'DOSAGEM DE COLESTEROL TOTAL',
            },
            {'codigo': '0202010317', 'descricao': 'DOSAGEM DE CREATININA'},
            {'codigo': '0202010473', 'descricao': 'DOSAGEM DE GLICOSE'},
            {
                'codigo': '0202010503',
                'descricao': 'DOSAGEM DE HEMOGLOBINA GLICOSILADA',
            },
            {'codigo': '0202010600', 'descricao': 'DOSAGEM DE POTASSIO'},
            {'codigo': '0202010635', 'descricao': 'DOSAGEM DE SODIO'},
            {
                'codigo': '0202010643',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-OXALACETICA (TGO)',
            },
            {
                'codigo': '0202010651',
                'descricao': 'DOSAGEM DE TRANSAMINASE GLUTAMICO-PIRUVICA (TGP)',
            },
            {'codigo': '0202010678', 'descricao': 'DOSAGEM DE TRIGLICERIDEOS'},
            {'codigo': '0202010694', 'descricao': 'DOSAGEM DE UREIA'},
            {'codigo': '0202020380', 'descricao': 'HEMOGRAMA COMPLETO'},
            {
                'codigo': '0205010032',
                'descricao': 'ECOCARDIOGRAFIA TRANSTORACICA',
            },
        ],
    },
    '0903010011': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA EM ORTOPEDIA COM RECURSOS DE RADIOLOGIA',
        'itens_obrigatorios': [
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0204020034',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO + OBLIQUAS)',
            },
            {
                'codigo': '0204020042',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO / FLEXAO)',
            },
            {
                'codigo': '0204020077',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA (C/ OBLIQUAS)',
            },
            {
                'codigo': '0204020085',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA FUNCIONAL / DINAMICA',
            },
            {
                'codigo': '0204020093',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACICA (AP + LATERAL)',
            },
            {
                'codigo': '0204020107',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACO-LOMBAR',
            },
            {
                'codigo': '0204020131',
                'descricao': 'RADIOGRAFIA PANORAMICA DE COLUNA TOTAL- TELESPONDILOGRAFIA ( P/ ESCOLIOSE)',
            },
            {
                'codigo': '0204040035',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO ESCAPULO-UMERAL',
            },
            {'codigo': '0204040078', 'descricao': 'RADIOGRAFIA DE COTOVELO'},
            {'codigo': '0204040094', 'descricao': 'RADIOGRAFIA DE MAO'},
            {
                'codigo': '0204040116',
                'descricao': 'RADIOGRAFIA DE ESCAPULA/OMBRO (TRES POSICOES)',
            },
            {
                'codigo': '0204040124',
                'descricao': 'RADIOGRAFIA DE PUNHO (AP + LATERAL + OBLIQUA)',
            },
            {
                'codigo': '0204060060',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO COXO-FEMORAL',
            },
            {'codigo': '0204060095', 'descricao': 'RADIOGRAFIA DE BACIA'},
            {'codigo': '0204060109', 'descricao': 'RADIOGRAFIA DE CALCANEO'},
            {
                'codigo': '0204060125',
                'descricao': 'RADIOGRAFIA DE JOELHO (AP + LATERAL)',
            },
            {
                'codigo': '0204060133',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + AXIAL)',
            },
            {
                'codigo': '0204060141',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + OBLIQUA + 3 AXIAIS)',
            },
            {
                'codigo': '0204060150',
                'descricao': 'RADIOGRAFIA DE PE / DEDOS DO PE',
            },
            {
                'codigo': '0204060176',
                'descricao': 'RADIOGRAFIA PANORAMICA DE MEMBROS INFERIORES',
            },
        ],
    },
    '0903010020': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA EM ORTOPEDIA COM RECURSOS DE RADIOLOGIA E ULTRASSONOGRAFIA',
        'itens_obrigatorios': [
            {
                'codigo': '0205020062',
                'descricao': 'ULTRASSONOGRAFIA DE ARTICULACAO',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0204020034',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO + OBLIQUAS)',
            },
            {
                'codigo': '0204020042',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO / FLEXAO)',
            },
            {
                'codigo': '0204020077',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA (C/ OBLIQUAS)',
            },
            {
                'codigo': '0204020085',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA FUNCIONAL / DINAMICA',
            },
            {
                'codigo': '0204020093',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACICA (AP + LATERAL)',
            },
            {
                'codigo': '0204020107',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACO-LOMBAR',
            },
            {
                'codigo': '0204020131',
                'descricao': 'RADIOGRAFIA PANORAMICA DE COLUNA TOTAL- TELESPONDILOGRAFIA ( P/ ESCOLIOSE)',
            },
            {
                'codigo': '0204040035',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO ESCAPULO-UMERAL',
            },
            {'codigo': '0204040078', 'descricao': 'RADIOGRAFIA DE COTOVELO'},
            {'codigo': '0204040094', 'descricao': 'RADIOGRAFIA DE MAO'},
            {
                'codigo': '0204040116',
                'descricao': 'RADIOGRAFIA DE ESCAPULA/OMBRO (TRES POSICOES)',
            },
            {
                'codigo': '0204040124',
                'descricao': 'RADIOGRAFIA DE PUNHO (AP + LATERAL + OBLIQUA)',
            },
            {
                'codigo': '0204060060',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO COXO-FEMORAL',
            },
            {'codigo': '0204060095', 'descricao': 'RADIOGRAFIA DE BACIA'},
            {'codigo': '0204060109', 'descricao': 'RADIOGRAFIA DE CALCANEO'},
            {
                'codigo': '0204060125',
                'descricao': 'RADIOGRAFIA DE JOELHO (AP + LATERAL)',
            },
            {
                'codigo': '0204060133',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + AXIAL)',
            },
            {
                'codigo': '0204060141',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + OBLIQUA + 3 AXIAIS)',
            },
            {
                'codigo': '0204060150',
                'descricao': 'RADIOGRAFIA DE PE / DEDOS DO PE',
            },
            {
                'codigo': '0204060176',
                'descricao': 'RADIOGRAFIA PANORAMICA DE MEMBROS INFERIORES',
            },
        ],
    },
    '0903010046': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA EM ORTOPEDIA COM RECURSOS DE RADIOLOGIA E RESSONÂNCIA MAGNÉTICA',
        'itens_obrigatorios': [
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
            {
                'codigo': '0301010307',
                'descricao': 'TELECONSULTA MÉDICA NA ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0204020034',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO + OBLIQUAS)',
            },
            {
                'codigo': '0204020042',
                'descricao': 'RADIOGRAFIA DE COLUNA CERVICAL (AP + LATERAL + TO / FLEXAO)',
            },
            {
                'codigo': '0204020077',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA (C/ OBLIQUAS)',
            },
            {
                'codigo': '0204020085',
                'descricao': 'RADIOGRAFIA DE COLUNA LOMBO-SACRA FUNCIONAL / DINAMICA',
            },
            {
                'codigo': '0204020093',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACICA (AP + LATERAL)',
            },
            {
                'codigo': '0204020107',
                'descricao': 'RADIOGRAFIA DE COLUNA TORACO-LOMBAR',
            },
            {
                'codigo': '0204020131',
                'descricao': 'RADIOGRAFIA PANORAMICA DE COLUNA TOTAL- TELESPONDILOGRAFIA ( P/ ESCOLIOSE)',
            },
            {
                'codigo': '0204040035',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO ESCAPULO-UMERAL',
            },
            {'codigo': '0204040078', 'descricao': 'RADIOGRAFIA DE COTOVELO'},
            {'codigo': '0204040094', 'descricao': 'RADIOGRAFIA DE MAO'},
            {
                'codigo': '0204040116',
                'descricao': 'RADIOGRAFIA DE ESCAPULA/OMBRO (TRES POSICOES)',
            },
            {
                'codigo': '0204040124',
                'descricao': 'RADIOGRAFIA DE PUNHO (AP + LATERAL + OBLIQUA)',
            },
            {
                'codigo': '0204060060',
                'descricao': 'RADIOGRAFIA DE ARTICULACAO COXO-FEMORAL',
            },
            {'codigo': '0204060095', 'descricao': 'RADIOGRAFIA DE BACIA'},
            {'codigo': '0204060109', 'descricao': 'RADIOGRAFIA DE CALCANEO'},
            {
                'codigo': '0204060125',
                'descricao': 'RADIOGRAFIA DE JOELHO (AP + LATERAL)',
            },
            {
                'codigo': '0204060133',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + AXIAL)',
            },
            {
                'codigo': '0204060141',
                'descricao': 'RADIOGRAFIA DE JOELHO OU PATELA (AP + LATERAL + OBLIQUA + 3 AXIAIS)',
            },
            {
                'codigo': '0204060150',
                'descricao': 'RADIOGRAFIA DE PE / DEDOS DO PE',
            },
            {
                'codigo': '0204060176',
                'descricao': 'RADIOGRAFIA PANORAMICA DE MEMBROS INFERIORES',
            },
            {
                'codigo': '0207010030',
                'descricao': 'RESSONANCIA MAGNETICA DE COLUNA CERVICAL/PESCOÇO',
            },
            {
                'codigo': '0207010048',
                'descricao': 'RESSONANCIA MAGNETICA DE COLUNA LOMBO-SACRA',
            },
            {
                'codigo': '0207010056',
                'descricao': 'RESSONANCIA MAGNETICA DE COLUNA TORACICA',
            },
            {
                'codigo': '0207020027',
                'descricao': 'RESSONANCIA MAGNETICA DE MEMBRO SUPERIOR (UNILATERAL)',
            },
            {
                'codigo': '0207030022',
                'descricao': 'RESSONANCIA MAGNETICA DE BACIA / PELVE / ABDOMEN INFERIOR',
            },
            {
                'codigo': '0207030030',
                'descricao': 'RESSONANCIA MAGNETICA DE MEMBRO INFERIOR (UNILATERAL)',
            },
        ],
    },
    '0904010015': {
        'nome': 'OCI AVALIAÇÃO INICIAL DIAGNÓSTICA DE DÉFICIT AUDITIVO',
        'itens_obrigatorios': [
            {
                'codigo': '0211070041',
                'descricao': 'AUDIOMETRIA TONAL LIMIAR (VIA AEREA / OSSEA)',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211070203', 'descricao': 'IMITANCIOMETRIA'}
        ],
    },
    '0904010023': {
        'nome': 'OCI PROGRESSÃO DA AVALIAÇÃO DIAGNÓSTICA DE DÉFICIT AUDITIVO',
        'itens_obrigatorios': [
            {
                'codigo': '0211070041',
                'descricao': 'AUDIOMETRIA TONAL LIMIAR (VIA AEREA / OSSEA)',
            },
            {
                'codigo': '0211070262',
                'descricao': 'POTENCIAL EVOCADO AUDITIVO DE CURTA MEDIA E LONGA LATENCIA',
            },
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0211050113',
                'descricao': 'POTENCIAL EVOCADO AUDITIVO',
            },
            {'codigo': '0211070203', 'descricao': 'IMITANCIOMETRIA'},
        ],
    },
    '0904010031': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA DE NASOFARINGE E DE OROFARINGE',
        'itens_obrigatorios': [
            {'codigo': '0209040025', 'descricao': 'LARINGOSCOPIA'},
            {'codigo': '0209040041', 'descricao': 'VIDEOLARINGOSCOPIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0905010019': {
        'nome': 'OCI AVALIAÇÃO INICIAL EM OFTALMOGIA - 0 A 8 ANOS',
        'itens_obrigatorios': [
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {'codigo': '0211060232', 'descricao': 'TESTE ORTÓPTICO'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0905010027': {
        'nome': 'OCI AVALIAÇÃO DE ESTRABISMO',
        'itens_obrigatorios': [
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {'codigo': '0211060232', 'descricao': 'TESTE ORTÓPTICO'},
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211060100', 'descricao': 'FUNDOSCOPIA'},
            {
                'codigo': '0211060178',
                'descricao': 'RETINOGRAFIA COLORIDA BINOCULAR',
            },
        ],
    },
    '0905010035': {
        'nome': 'OCI AVALIAÇÃO INICIAL EM OFTALMOLOGIA - A PARTIR DE 9 ANOS',
        'itens_obrigatorios': [
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211060232', 'descricao': 'TESTE ORTÓPTICO'}
        ],
    },
    '0905010043': {
        'nome': 'OCI AVALIAÇÃO DE RETINOPATIA DIABÉTICA',
        'itens_obrigatorios': [
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {
                'codigo': '0211060178',
                'descricao': 'RETINOGRAFIA COLORIDA BINOCULAR',
            },
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0905010051': {
        'nome': 'OCI AVALIAÇÃO INICIAL PARA ONCOLOGIA OFTALMOLÓGICA',
        'itens_obrigatorios': [
            {
                'codigo': '0205020089',
                'descricao': 'ULTRASSONOGRAFIA DE GLOBO OCULAR / ORBITA (MONOCULAR)',
            },
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {
                'codigo': '0211060178',
                'descricao': 'RETINOGRAFIA COLORIDA BINOCULAR',
            }
        ],
    },
    '0905010060': {
        'nome': 'OCI AVALIAÇÃO DIAGNÓSTICA EM NEURO OFTALMOLOGIA',
        'itens_obrigatorios': [
            {
                'codigo': '0211060020',
                'descricao': 'BIOMICROSCOPIA DE FUNDO DE OLHO',
            },
            {
                'codigo': '0211060038',
                'descricao': 'CAMPIMETRIA COMPUTADORIZADA OU MANUAL COM GRÁFICO',
            },
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {
                'codigo': '0211060178',
                'descricao': 'RETINOGRAFIA COLORIDA BINOCULAR',
            },
            {'codigo': '0211060224', 'descricao': 'TESTE DE VISÃO DE CORES'},
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [],
    },
    '0905010078': {
        'nome': 'OCI EXAMES OFTALMOLÓGICOS SOB SEDAÇÃO',
        'itens_obrigatorios': [
            {'codigo': '0417010060', 'descricao': 'SEDACAO'},
            {
                'codigo': '0301010072',
                'descricao': 'CONSULTA MEDICA EM ATENÇÃO ESPECIALIZADA',
            },
        ],
        'itens_facultativos': [
            {'codigo': '0211060127', 'descricao': 'MAPEAMENTO DE RETINA'},
            {'codigo': '0211060259', 'descricao': 'TONOMETRIA'},
        ],
    },
}


def allowed_file(filename):
    return ('.' in filename and 
            filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS'])

def formatar_dados(df):
    # Função auxiliar para aplicar zfill apenas se for número
    def aplicar_zfill(serie, tamanho):
        serie = serie.fillna('').astype(str)
        return serie.where(
            ~serie.str.match(r'^\d+$'), serie.str.zfill(tamanho)
        )

    # CNES - 7 dígitos
    cnes_cols = ['CNES_SOLICITANTE', 'CNES_REGULADOR', 'CNES_EXECUTANTE']
    for col in cnes_cols:
        if col in df.columns:
            df[col] = aplicar_zfill(df[col], 7)

    # CODIGO_SIGTAP - 10 dígitos
    if 'CODIGO_SIGTAP' in df.columns:
        df['CODIGO_SIGTAP'] = aplicar_zfill(df['CODIGO_SIGTAP'], 10)

    # Formatação de datas
    date_cols = ['DATA_SOLICITACAO', 'DATA_AUTORIZACAO', 'DATA_EXECUCAO']
    for col in date_cols:
        if col in df.columns:
            # Primeiro tenta converter para datetime
            df[col] = pd.to_datetime(df[col], errors='coerce', format='%Y-%m-%d')
            
            # Para valores que não converteram, tenta outros formatos
            mask = df[col].isna()
            if mask.any():
                df.loc[mask, col] = pd.to_datetime(
                    df.loc[mask, col], 
                    errors='coerce',
                    format='%d/%m/%Y'
                )
            
            # Formata para string no formato desejado
            df[col] = df[col].dt.strftime('%d/%m/%Y')
            df[col] = df[col].replace('NaT', '')

    # Tratamento especial para CBO - converter NaN para string vazia
    if 'CBO' in df.columns:
        df['CBO'] = df['CBO'].fillna('').astype(str)
    
    return df

def analisar_dados(df):
    try:
        # Filtra apenas registros com STATUS == 1 (Em Espera)
        df = df[df['STATUS'] == '1'].copy()
        
        # Preenche NaN com string vazia
        df = df.fillna('')

        total_solicitacoes = len(df)
        relatorio = []
        relatorio_agrupamentos = []  # Lista para dados XLSX de agrupamentos
        relatorio_nao_agrupados = []  # Lista para dados XLSX de não agrupados

        relatorio.append(
            "*********************    FORAM ENCONTRADOS {} CONJUNTOS DE OCI'S    ***********************\n"
        )

        total_pacientes = len(df['DOCUMENTO_PACIENTE'].unique())
        pacientes_em_agrupamentos = set()
        agrupamentos_encontrados = 0

        for codigo, agrupamento in agrupamentos.items():
            codigos_obrigatorios = [
                item['codigo'] for item in agrupamento['itens_obrigatorios']
            ]
            codigos_facultativos = [
                item['codigo'] for item in agrupamento['itens_facultativos']
            ]

            pacientes_agrupados = []

            for paciente, dados_paciente in df.groupby('DOCUMENTO_PACIENTE'):
                codigos_paciente = set(dados_paciente['CODIGO_SIGTAP'])

                if all(c in codigos_paciente for c in codigos_obrigatorios):
                    pacientes_agrupados.append(paciente)
                    pacientes_em_agrupamentos.add(paciente)

            if pacientes_agrupados:
                agrupamentos_encontrados += 1
                relatorio.append(
                    '_________________________________________________________________________________________'
                )
                relatorio.append(
                    f"{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]} - {agrupamento['nome']}\n"
                )

                for paciente in sorted(pacientes_agrupados):
                    relatorio.append(f'--- {paciente}')

                    # Itens obrigatórios
                    for item in agrupamento['itens_obrigatorios']:
                        registros = df[
                            (df['DOCUMENTO_PACIENTE'] == paciente)
                            & (df['CODIGO_SIGTAP'] == item['codigo'])
                        ]
                        for _, registro in registros.iterrows():
                            relatorio.append(
                                f"-------- OBG\tCNES_SOLC {registro['CNES_SOLICITANTE']}\tCID-{registro['CID10']}\tDT_SOLC-{registro['DATA_SOLICITACAO']}\t{item['codigo']} - {item['descricao']}"
                            )
                            relatorio_agrupamentos.append({
                                'AGRUPAMENTO_OCI': f'{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}',
                                'DESCRICAO_OCI': agrupamento['nome'],
                                'DOCUMENTO_PACIENTE': paciente,
                                'DATA_SOLICITACAO': registro['DATA_SOLICITACAO'],
                                'CNES_SOLICITANTE': registro['CNES_SOLICITANTE'],
                                'ITEM OBG/FAC (X)': 'OBG',
                                'CID10': registro['CID10'],
                                'CODIGO_SIGTAP': item['codigo'],
                                'DESCRICAO_SIGTAP': item['descricao'],
                                'CBO': registro['CBO']
                            })

                    # Itens facultativos
                    for item in agrupamento['itens_facultativos']:
                        registros = df[
                            (df['DOCUMENTO_PACIENTE'] == paciente)
                            & (df['CODIGO_SIGTAP'] == item['codigo'])
                        ]
                        for _, registro in registros.iterrows():
                            relatorio.append(
                                f"-------- FAC\tCNES_SOLC {registro['CNES_SOLICITANTE']}\tCID-{registro['CID10']}\tDT_SOLC-{registro['DATA_SOLICITACAO']}\t{item['codigo']} - {item['descricao']}"
                            )
                            relatorio_agrupamentos.append({
                                'AGRUPAMENTO_OCI': f'{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}',
                                'DESCRICAO_OCI': agrupamento['nome'],
                                'DOCUMENTO_PACIENTE': paciente,
                                'DATA_SOLICITACAO': registro['DATA_SOLICITACAO'],
                                'CNES_SOLICITANTE': registro['CNES_SOLICITANTE'],
                                'ITEM OBG/FAC (X)': 'FAC',
                                'CID10': registro['CID10'],
                                'CODIGO_SIGTAP': item['codigo'],
                                'DESCRICAO_SIGTAP': item['descricao'],
                                'CBO': registro['CBO']
                            })

        # Pacientes não agrupados
        pacientes_restantes = df[~df['DOCUMENTO_PACIENTE'].isin(pacientes_em_agrupamentos)]
        relatorio.append(
            '\n********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO  ***********************'
        )
        for _, linha in pacientes_restantes.iterrows():
            descricao = next(
                (item['descricao']
                for g in agrupamentos.values()
                for item in g['itens_obrigatorios'] + g['itens_facultativos']
                if item['codigo'] == linha['CODIGO_SIGTAP']
                ),
                'Código não faz parte de um item de OCI',
            )
            relatorio.append(
                f"- CNES_SOLC {linha['CNES_SOLICITANTE']}\tCID {linha['CID10']}\tCNS/CPF_PAC {linha['DOCUMENTO_PACIENTE']}\tDT_SOLC {linha['DATA_SOLICITACAO']}\t{linha['CODIGO_SIGTAP']} - {descricao}"
            )
            relatorio_nao_agrupados.append({
                'DOCUMENTO_PACIENTE': linha['DOCUMENTO_PACIENTE'],
                'DATA_SOLICITACAO': linha['DATA_SOLICITACAO'],
                'CNES_SOLICITANTE': linha['CNES_SOLICITANTE'],
                'CID10': linha['CID10'],
                'CODIGO_SIGTAP': linha['CODIGO_SIGTAP'],
                'DESCRICAO_SIGTAP': descricao,
                'CBO': linha['CBO']
            })

        # Atualiza o cabeçalho com o número real de conjuntos
        relatorio[0] = relatorio[0].format(agrupamentos_encontrados)

        # Rodapé com data e hora
        relatorio.append('\n{:%d/%m/%Y %H:%M:%S}'.format(datetime.now()))

        return {
            'relatorio': relatorio,
            'relatorio_agrupamentos': relatorio_agrupamentos,
            'relatorio_nao_agrupados': relatorio_nao_agrupados,
            'total_pacientes': total_pacientes,
            'pacientes_agrupados': len(pacientes_em_agrupamentos),
            'agrupamentos_encontrados': agrupamentos_encontrados,
            'total_solicitacoes': total_solicitacoes,
        }

    except Exception as e:
        app.logger.error(f"Erro na análise de dados: {str(e)}")
        raise e

def gerar_pdf(relatorio, tempo_processamento):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    width, height = landscape(letter)
    
    # Configurações de layout
    margem_esquerda = 30
    linha_altura = 12
    max_linhas_por_pagina = 50
    pagina_numero = 1
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    
    # Filtra apenas as linhas de agrupamentos
    linhas_agrupamentos = []
    for linha in relatorio:
        if linha.startswith('********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO'):
            break
        linhas_agrupamentos.append(linha)
    
    # Função para desenhar cabeçalho
    def desenhar_cabecalho(pagina):
        try:
            c.drawImage('static/img/agora-tem-especialistas.png', margem_esquerda, height - 80, width=100, height=50)
        except:
            pass  # Ignora erro se imagem não existir
        
        c.setFont('Helvetica-Bold', 16)
        c.drawString(margem_esquerda + 120, height - 60, "Relatório de Análise de Filas OCI")
        c.setFont('Helvetica', 10)
        c.drawString(margem_esquerda + 120, height - 80, f"Gerado em: {data_hora}")
        c.drawString(margem_esquerda + 120, height - 95, f"Tempo de processamento: {tempo_processamento} segundos")
        c.drawRightString(width - margem_esquerda, height - 95, f"Página {pagina}")
        c.line(margem_esquerda, height - 105, width - margem_esquerda, height - 105)
    
    # Processamento das linhas
    linha_index = 0
    total_linhas = len(linhas_agrupamentos)
    
    while linha_index < total_linhas:
        # Configurar nova página
        c.setPageSize(landscape(letter))
        desenhar_cabecalho(pagina_numero)
        y = height - 120  # Posição inicial do conteúdo
        
        # Adicionar conteúdo da página
        linhas_na_pagina = 0
        while linhas_na_pagina < max_linhas_por_pagina and linha_index < total_linhas:
            linha = linhas_agrupamentos[linha_index]
            
            # Escolher fonte apropriada
            if linha.startswith('---') or linha.startswith('--------'):
                c.setFont('Courier', 8)
            else:
                c.setFont('Helvetica-Bold', 9)
                
            c.drawString(margem_esquerda, y, linha[:200])  # Limita para evitar overflow
            
            y -= linha_altura
            linha_index += 1
            linhas_na_pagina += 1
            
            # Verificar se chegou ao final da página
            if y < 50:
                break
        
        # Finalizar página e preparar próxima
        c.showPage()
        pagina_numero += 1
    
    # Salvar PDF
    c.save()
    buffer.seek(0)
    return buffer

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze_file', methods=['POST'])
def analyze_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo vazio'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Tipo de arquivo não permitido. Use .csv ou .xlsx'}), 400

    try:
        tempo_inicio = time()

        # Lê o arquivo conforme o tipo
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, encoding='utf-8', sep=';', dtype=str)
        else:  # XLSX
            df = pd.read_excel(file, dtype=str)

        # Preenche valores NaN com string vazia
        df = df.fillna('')

        tempo_leitura = time()

        # Aplica as formatações necessárias
        df = formatar_dados(df)
        tempo_formatacao = time()

        # Verifica colunas obrigatórias
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            return jsonify({
                'message': 'Colunas obrigatórias faltando no arquivo',
                'details': {'missing_columns': missing_columns},
            }), 400

        # Realiza a análise
        resultado = analisar_dados(df)
        tempo_analise = time()

        tempo_total = tempo_analise - tempo_inicio

        return jsonify({
            'success': True,
            'relatorio': resultado['relatorio'],
            'relatorio_agrupamentos': resultado['relatorio_agrupamentos'],
            'relatorio_nao_agrupados': resultado['relatorio_nao_agrupados'],
            'resumo': {
                'total_pacientes': resultado['total_pacientes'],
                'total_solicitacoes': resultado['total_solicitacoes'],
                'pacientes_agrupados': resultado['pacientes_agrupados'],
                'agrupamentos_encontrados': resultado['agrupamentos_encontrados'],
                'tempo_processamento': round(tempo_total, 2),
                'tempos_parciais': {
                    'leitura': round(tempo_leitura - tempo_inicio, 2),
                    'formatacao': round(tempo_formatacao - tempo_leitura, 2),
                    'analise': round(tempo_analise - tempo_formatacao, 2),
                },
            }
        })

    except Exception as e:
        app.logger.error(f"Erro ao processar arquivo: {str(e)}")
        return jsonify({'error': f'Erro ao processar arquivo: {str(e)}'}), 500

@app.route('/download_pdf', methods=['POST'])
def download_pdf():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Dados não fornecidos'}), 400

        relatorio = data.get('relatorio')
        tempo_processamento = data.get('tempo_processamento')
        
        if not relatorio or tempo_processamento is None:
            return jsonify({'error': 'Parâmetros faltando'}), 400

        pdf_buffer = gerar_pdf(relatorio, tempo_processamento)
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name='relatorio_agrupamentos_oci.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        app.logger.error(f"Erro ao gerar PDF: {str(e)}")
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500

@app.route('/download_xlsx', methods=['POST'])
def download_xlsx():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Dados não fornecidos'}), 400

        relatorio_agrupamentos = data.get('relatorio_agrupamentos', [])
        relatorio_nao_agrupados = data.get('relatorio_nao_agrupados', [])
        
        if not relatorio_agrupamentos and not relatorio_nao_agrupados:
            return jsonify({'error': 'Nenhum relatório fornecido'}), 400

        # Cria DataFrames - tratamento seguro para dados ausentes
        df_agrupamentos = pd.DataFrame(relatorio_agrupamentos) if relatorio_agrupamentos else pd.DataFrame()
        df_nao_agrupados = pd.DataFrame(relatorio_nao_agrupados) if relatorio_nao_agrupados else pd.DataFrame()
        
        # Cria buffer Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Aba de agrupamentos
            if not df_agrupamentos.empty:
                df_agrupamentos.to_excel(writer, sheet_name='Agrupamentos', index=False)
            
            # Aba de não agrupados
            if not df_nao_agrupados.empty:
                df_nao_agrupados.to_excel(writer, sheet_name='NaoAgrupados', index=False)
            
            # Formatação
            workbook = writer.book
            
            # Formatação para cabeçalhos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            # Aplica formatação para cada aba
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                
                # Determina colunas com base na aba
                if sheet_name == 'Agrupamentos':
                    columns = df_agrupamentos.columns
                else:
                    columns = df_nao_agrupados.columns
                
                # Aplica formatação ao cabeçalho
                for col_num, value in enumerate(columns):
                    worksheet.write(0, col_num, value, header_format)
                
                # Ajusta largura das colunas
                for col_num, column in enumerate(columns):
                    max_len = 0
                    if sheet_name == 'Agrupamentos':
                        max_len = df_agrupamentos[column].astype(str).map(len).max()
                    else:
                        max_len = df_nao_agrupados[column].astype(str).map(len).max()
                    
                    if pd.isna(max_len):
                        max_len = len(column)
                    else:
                        max_len = max(max_len, len(column))
                    
                    worksheet.set_column(col_num, col_num, min(max_len + 2, 50))
                
                # Congela cabeçalho
                worksheet.freeze_panes(1, 0)

        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f'relatorio_oci_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        app.logger.error(f"Erro ao gerar XLSX: {str(e)}")
        return jsonify({'error': f'Erro ao gerar XLSX: {str(e)}'}), 500

@app.route('/download-modelo')
def download_modelo():
    return send_from_directory('DB', 'arquivo_modelo.xlsx', as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)