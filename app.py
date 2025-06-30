import io
import os
import re
import tempfile
from datetime import datetime
from time import time

import pandas as pd
from flask import (Flask, jsonify, render_template, request, send_file,
                   send_from_directory)
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

def detectar_formato_data(data_str):
    """Detecta o formato da data em uma string"""
    if pd.isna(data_str) or data_str.strip() == '':
        return None

    padroes = [
        (r'^\d{4}-\d{2}-\d{2}$', '%Y-%m-%d'),
        (r'^\d{4}/\d{2}/\d{2}$', '%Y/%m/%d'),
        (r'^\d{2}/\d{2}/\d{4}$', '%d/%m/%Y'),
        (r'^\d{2}-\d{2}-\d{4}$', '%d-%m-%Y'),
        (r'^\d{2}/\d{2}/\d{2}$', '%d/%m/%y'),
        (r'^\d{2}-\d{2}-\d{2}$', '%d-%m-%y'),
        (r'^\d{8}$', '%d%m%Y'),
        (r'^\d{1,2}/\d{1,2}/\d{4}$', '%d/%m/%Y'),
        (r'^\d{1,2}-\d{1,2}-\d{4}$', '%d-%m-%Y'),
    ]

    for padrao, formato in padroes:
        if re.match(padrao, data_str.strip()):
            return formato
    return None

def converter_data(data_str):
    """Converte uma string de data para o formato AAAA-MM-DD"""
    if pd.isna(data_str) or data_str.strip() == '':
        return ''

    formato = detectar_formato_data(data_str)
    if not formato:
        return data_str

    try:
        data_obj = datetime.strptime(data_str.strip(), formato)
        return data_obj.strftime('%Y-%m-%d')
    except ValueError:
        return data_str

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
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
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
        relatorio_xlsx = []
        pacientes_em_agrupamentos = set()
        agrupamentos_encontrados = 0

        relatorio.append(
            "*********************    FORAM ENCONTRADOS {} CONJUNTOS DE OCI'S    ***********************\n"
        )

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
                            relatorio_xlsx.append({
                                'AGRUPAMENTO_OCI': f'{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}',
                                'DESCRICAO_OCI': agrupamento['nome'],
                                'DOCUMENTO_PACIENTE': paciente,
                                'DATA_SOLICITACAO': registro['DATA_SOLICITACAO'],
                                'CNES_SOLICITANTE': registro['CNES_SOLICITANTE'],
                                'TIPO_ITEM': 'OBRIGATÓRIO',
                                'CID10': registro['CID10'],
                                'CODIGO_SIGTAP': item['codigo'],
                                'DESCRICAO_SIGTAP': item['descricao'],
                                'CBO': registro['CBO'],
                                'CNES_EXECUTANTE': registro.get('CNES_EXECUTANTE', ''),
                                'DATA_AUTORIZACAO': registro.get('DATA_AUTORIZACAO', ''),
                                'DATA_EXECUCAO': registro.get('DATA_EXECUCAO', ''),
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
                            relatorio_xlsx.append({
                                'AGRUPAMENTO_OCI': f'{codigo[:2]}.{codigo[2:4]}.{codigo[4:6]}.{codigo[6:]}',
                                'DESCRICAO_OCI': agrupamento['nome'],
                                'DOCUMENTO_PACIENTE': paciente,
                                'DATA_SOLICITACAO': registro['DATA_SOLICITACAO'],
                                'CNES_SOLICITANTE': registro['CNES_SOLICITANTE'],
                                'TIPO_ITEM': 'FACULTATIVO',
                                'CID10': registro['CID10'],
                                'CODIGO_SIGTAP': item['codigo'],
                                'DESCRICAO_SIGTAP': item['descricao'],
                                'CBO': registro['CBO'],
                                'CNES_EXECUTANTE': registro.get('CNES_EXECUTANTE', ''),
                                'DATA_AUTORIZACAO': registro.get('DATA_AUTORIZACAO', ''),
                                'DATA_EXECUCAO': registro.get('DATA_EXECUCAO', ''),
                            })

        # Adiciona pacientes não agrupados ao relatório XLSX
        pacientes_restantes = df[~df['DOCUMENTO_PACIENTE'].isin(pacientes_em_agrupamentos)]
        for _, registro in pacientes_restantes.iterrows():
            descricao = next(
                (item['descricao']
                for g in agrupamentos.values()
                for item in g['itens_obrigatorios'] + g['itens_facultativos']
                if item['codigo'] == registro['CODIGO_SIGTAP']
                ),
                'Código não faz parte de um item de OCI',
            )
            relatorio_xlsx.append({
                'AGRUPAMENTO_OCI': 'NÃO AGRUPADO',
                'DESCRICAO_OCI': 'NÃO AGRUPADO',
                'DOCUMENTO_PACIENTE': registro['DOCUMENTO_PACIENTE'],
                'DATA_SOLICITACAO': registro['DATA_SOLICITACAO'],
                'CNES_SOLICITANTE': registro['CNES_SOLICITANTE'],
                'TIPO_ITEM': 'NÃO AGRUPADO',
                'CID10': registro['CID10'],
                'CODIGO_SIGTAP': registro['CODIGO_SIGTAP'],
                'DESCRICAO_SIGTAP': descricao,
                'CBO': registro['CBO'],
                'CNES_EXECUTANTE': registro.get('CNES_EXECUTANTE', ''),
                'DATA_AUTORIZACAO': registro.get('DATA_AUTORIZACAO', ''),
                'DATA_EXECUCAO': registro.get('DATA_EXECUCAO', ''),
            })

        # Atualiza o cabeçalho com o número real de conjuntos
        relatorio[0] = relatorio[0].format(agrupamentos_encontrados)

        # Pacientes que não estão em nenhum agrupamento (para relatório textual)
        relatorio.append(
            '\n********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO  ***********************'
        )
        for _, linha in pacientes_restantes.iterrows():
            data_solc = linha['DATA_SOLICITACAO'].split(' ')[0] if linha['DATA_SOLICITACAO'] else ''
            relatorio.append(
                f"- CNES_SOLC {linha['CNES_SOLICITANTE']}\tCID {linha['CID10']}\tCNS/CPF_PAC {linha['DOCUMENTO_PACIENTE']}\tDT_SOLC {data_solc}\t{linha['CODIGO_SIGTAP']} - {descricao}"
            )

        # Rodapé com data e hora
        relatorio.append('\n{:%d/%m/%Y %H:%M:%S}'.format(datetime.now()))

        # Converte NaN para None em todo o relatório_xlsx
        relatorio_xlsx_sanitizado = []
        for item in relatorio_xlsx:
            sanitized_item = {k: (v if pd.notna(v) else '') for k, v in item.items()}
            relatorio_xlsx_sanitizado.append(sanitized_item)

        return {
            'relatorio': relatorio,
            'relatorio_xlsx': relatorio_xlsx_sanitizado,
            'total_pacientes': len(df['DOCUMENTO_PACIENTE'].unique()),
            'pacientes_agrupados': len(pacientes_em_agrupamentos),
            'agrupamentos_encontrados': agrupamentos_encontrados,
            'total_solicitacoes': total_solicitacoes,
        }

    except Exception as e:
        raise e

def gerar_pdf(relatorio, tempo_processamento):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    width, height = landscape(letter)

    # Cabeçalho com logo
    c.drawImage('static/img/agora-tem-especialistas.png', 30, height - 80, width=100, height=50)
    c.setFont('Helvetica-Bold', 16)
    c.drawString(150, height - 60, "Relatório de Análise de Filas OCI")
    c.setFont('Helvetica', 10)
    c.drawString(150, height - 80, f"Gerado em: {datetime.now().strftime('%d/%m/%Y')}")
    c.drawString(150, height - 95, f"Tempo de processamento: {tempo_processamento} segundos")
    c.line(30, height - 105, width - 30, height - 105)

    c.setFont('Courier', 8)
    y = height - 120
    margem_esquerda = 30

    # Filtra apenas as linhas que contêm os agrupamentos OCI
    linhas_agrupamentos = [linha for linha in relatorio if not linha.startswith('********************    PACIENTES QUE NÃO ESTÃO EM NENHUM CONJUNTO')]
    
    for linha in linhas_agrupamentos:
        # Remove a hora das datas
        linha_formatada = re.sub(r'(\d{4}-\d{2}-\d{2}) 00:00:00', r'\1', linha)
        c.drawString(margem_esquerda, y, linha_formatada[:300])
        y -= 10
        if y < 30:
            c.showPage()
            c.setFont('Courier', 8)
            y = height - 30

    # Rodapé
    c.setFont('Helvetica', 8)
    c.drawString(margem_esquerda, 20, f"VERSÃO 1.0 - Gerado em {datetime.now().strftime('%d/%m/%Y')}")
    
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
            'relatorio_xlsx': resultado['relatorio_xlsx'],
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
        return jsonify({'error': f'Erro ao processar arquivo: {str(e)}'}), 500

@app.route('/download_pdf', methods=['POST'])
def download_pdf():
    try:
        relatorio = request.json.get('relatorio')
        tempo_processamento = request.json.get('tempo_processamento')
        if not relatorio:
            return jsonify({'error': 'Nenhum relatório fornecido'}), 400

        pdf_buffer = gerar_pdf(relatorio, tempo_processamento)
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name='relatorio_agrupamentos_oci.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': f'Erro ao gerar PDF: {str(e)}'}), 500

@app.route('/download_xlsx', methods=['POST'])
def download_xlsx():
    try:
        relatorio_xlsx = request.json.get('relatorio_xlsx')
        if not relatorio_xlsx:
            return jsonify({'error': 'Nenhum relatório fornecido'}), 400

        # Cria o DataFrame com os dados
        df = pd.DataFrame(relatorio_xlsx)
        
        # Ordena os dados
        df.sort_values(
            by=['AGRUPAMENTO_OCI', 'DOCUMENTO_PACIENTE', 'TIPO_ITEM', 'CODIGO_SIGTAP'],
            inplace=True
        )
        
        # Cria um buffer para o arquivo Excel
        output = io.BytesIO()
        
        # Configura o ExcelWriter com formatação
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Relatorio_OCI', index=False)
            
            # Acessa a planilha e o workbook para formatação
            workbook = writer.book
            worksheet = writer.sheets['Relatorio_OCI']
            
            # Formatação para cabeçalhos
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            # Aplica formatação aos cabeçalhos
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Ajusta a largura das colunas
            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),  # Maior valor na coluna
                    len(col)  # Tamanho do cabeçalho
                )
                worksheet.set_column(i, i, min(max_len + 2, 50))  # Limita a 50 caracteres
            
            # Congela a primeira linha (cabeçalhos)
            worksheet.freeze_panes(1, 0)

        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f'relatorio_oci_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': f'Erro ao gerar XLSX: {str(e)}'}), 500

@app.route('/download-modelo')
def download_modelo():
    return send_from_directory('DB', 'arquivo_modelo.xlsx', as_attachment=True)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)