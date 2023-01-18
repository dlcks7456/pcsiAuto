import os
import olefile
import zlib
import struct
from collections import OrderedDict
import openpyxl
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment, Color, fills, Side
from copy import copy

def get_hwp_text(filename):
    f = olefile.OleFileIO(filename)
    dirs = f.listdir()

    # HWP 파일 검증
    if ["FileHeader"] not in dirs or \
       ["\x05HwpSummaryInformation"] not in dirs:
        raise Exception("Not Valid HWP.")

    # 문서 포맷 압축 여부 확인
    header = f.openstream("FileHeader")
    header_data = header.read()
    is_compressed = (header_data[36] & 1) == 1

    # Body Sections 불러오기
    nums = []
    for d in dirs:
        if d[0] == "BodyText":
            nums.append(int(d[1][len("Section"):]))
    sections = ["BodyText/Section"+str(x) for x in sorted(nums)]

    # 예외 처리 
    bad_bytes = [
        '\x0b漠杳\x00\x00\x00\x00\x0b',
        '\x0b氠瑢\x00\x00\x00\x00\x0b',
        '\x15湯湷\x00\x00\x00\x00\x15',
        '\U000f0288'
    ]

    # 전체 text 추출
    text = ""
    for section in sections:
        bodytext = f.openstream(section)
        data = bodytext.read()
        if is_compressed:
            unpacked_data = zlib.decompress(data, -15)
        else:
            unpacked_data = data
    
        # 각 Section 내 text 추출    
        section_text = ""
        i = 0
        size = len(unpacked_data)
        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in [67]:
                rec_data = unpacked_data[i+4:i+4+rec_len]
                decode_rec = rec_data.decode('utf-16')
                for bad in bad_bytes :
                    if bad in decode_rec :
                        decode_rec = ''

                if not decode_rec == '' : 
                    section_text += decode_rec
                    section_text += "\n"
                else :
                    section_text += '-'*16
                    section_text += "\r\n"

            i += 4 + rec_len

        text += section_text
        text += "\n"

    f.close()
    
    return text

def pcsi_setting(survey_name='', 
                division='', 
                info_text_key = '면접원 지시사항', 
                qnr_folder = 'QNR', 
                save_folder = 'SET') :

    if survey_name == '' or not type(survey_name) == str :
        print('❌ ERROR : 기관명은 문자형으로 입력')
        return

    if division == '' or not type(division) == str or not division in ['KMAC', 'KSA']:
        print('❌ ERROR : 구분은 KMAC/KSA로만 입력 (대소문자 정확하게)')
        return

    hwps = os.listdir(qnr_folder)
    hwps = [i for i in hwps if '.hwp' in i]

    key_texts = []
    if division == 'KMAC' :
        key_texts = ['SQ3', 'SQ4']

    if division == 'KSA' : 
        key_texts = ['SQ3', 'SQ4', '문 9', '문 10', '문 12']
    # 워딩이 다른 문항
    type_code = {
        'A' : 1,
        'B' : 2,
        'C' : 3,
        'D' : 4,
        'E' : 5,
        'F' : 6,
        'G' : 7,
        'H' : 8
    }

    change_cells = OrderedDict()
    for key in key_texts :
        change_cells[key] = OrderedDict()

    change_cells['qnrs'] = OrderedDict()
    change_cells['info'] = OrderedDict()
    change_cells['Q8'] = OrderedDict()


    for hwp in hwps :
        # QNR 세팅
        del_hwp = hwp.replace('.hwp', '')
        code, label = del_hwp.split('.')
        name, qtype = label.split('_')
        change_cells['qnrs'][code] = {'name': name, 'type': qtype, 'type_code': type_code[qtype]}
        
        # SQ 세팅 (공통)
        curr_hwp = get_hwp_text(os.path.join(os.getcwd(), qnr_folder, hwp)).split('\r\n')
        for key in key_texts :
            curr_txt = [i for i in curr_hwp if key in i]
            if not curr_txt :
                continue
            set_word = curr_txt[0]
            set_word = set_word.replace('○○', '고객')
            set_word = set_word.replace(f'{key}. ', '')
            set_word = set_word.replace(f'{key}】 ', '')
            if name in set_word :
                set_word = set_word.replace(survey_name, f'<font color=blue>{survey_name}</font>')
            change_cells[key][code] = set_word.strip()
        
        # KSA ONLY
        if division == 'KSA' :
            Q8_array = []
            for idx, tx in enumerate(curr_hwp) :
                if '8-' in tx and not '】' in tx:
                    Q8_array.append(curr_hwp[idx+1])

            change_cells['Q8'][code] = Q8_array

        # SQ 이 후 조사 시작전 안내 문구
        info_txt = []
        info_flag = False
        for tx in curr_hwp :
            if not info_text_key and not info_text_key in tx :
                continue
            if info_text_key in tx :
                info_flag = True
                continue
            
            if info_flag and '-'*16 == tx :
                break

            if info_flag :
                set_word = tx
                if survey_name in set_word :
                    set_word = set_word.replace(survey_name, f'<font color=blue>{survey_name}</font>')
                info_txt.append(set_word.strip())
        
        info_txt = '<br/><br/>'.join(info_txt)
        change_cells['info'][code] = info_txt
        

    wb = openpyxl.load_workbook('template.xlsx')
    wb_sheetname = wb.sheetnames[0]
    ws = wb[wb_sheetname]
    rows = ws.rows
    cols = ws.columns

    new_wb = openpyxl.Workbook()
    new_wb.active.title = wb_sheetname
    new_ws = new_wb.active

    for row in rows :
        for cell in row :
            curr_cell = new_ws.cell(row=cell.row, column=cell.column)
            curr_cell.value = cell.value
            if cell.has_style :
                curr_cell.font = copy(cell.font)
                curr_cell.border = copy(cell.border)
                curr_cell.fill = copy(cell.fill)
                curr_cell.number_format = copy(cell.number_format)
                curr_cell.protection = copy(cell.protection)
                curr_cell.alignment = copy(cell.alignment)

    use_columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
    for us in use_columns :
        new_ws.column_dimensions[us].width = ws.column_dimensions[us].width


    # 기관명 세팅 관련
    name_set = new_ws.cell(27, 6)
    name_set.value = name_set.value%(survey_name)


    # 설문지 분류 셀 관련
    cell_value = []
    for code, attr in change_cells['qnrs'].items() :
        qname = attr['name']
        word = f'({code}) {qname}'
        cell_value.append(word)


    set_cells = [(13, 6), (25, 6)]
    for r, c in set_cells :
        set_sell = new_ws.cell(r, c)
        set_sell.value = set_sell.value%('\n'.join(cell_value))

    # 기관명 & 설문 분류
    js_logics = []
    for code, attr in change_cells['qnrs'].items() :
        qname = attr['name']
        word = f'if(ADD05=={code}){{text=\'{qname}\';}}'
        js_logics.append(word)

    QQQ1_set = new_ws.cell(16, 8)
    QQQ1_set.value = QQQ1_set.value%(survey_name, '\n'.join(js_logics))

    # 설문 타입 오토펀치 syntax
    q_type_quto = []
    for code, attr in change_cells['qnrs'].items() :
        qtype = attr['type_code']
        word = f'if(QQQ14=={code}) then TQ1={qtype}'
        q_type_quto.append(word)

    TQ1_set = new_ws.cell(29, 7)
    TQ1_set.value = TQ1_set.value%('\n'.join(q_type_quto))

    # 워딩 다른 문항 출력
    Q_cell_dict = {
        'SQ3' : (30, 8),
        'SQ4' : (31, 8),
        '문 9' : (61, 8),
        '문 10' : (62, 8), 
        '문 12' : (64, 8),
    }

    for qid in key_texts :
        cr, cc = Q_cell_dict[qid]
        curr_cell = new_ws.cell(cr, cc)
        js_logics = []
        for code, txt in change_cells[qid].items() :
            word = f'if(QQQ14=={code}){{text=\'{txt}\';}}'
            js_logics.append(word)
        
        curr_cell.value = curr_cell.value%('\n'.join(js_logics))


    # SQ 문항 이후 안내 문구 출력
    info_texts = []
    for code, txt in change_cells['info'].items() :
        word = f'if(QQQ14=={code}){{text=\'{txt}\';}}'
        info_texts.append(word)


    # QQQ1/Q1/Q5/Q6 워딩 구분
    QQQ1_txt = {
        'KMAC' : '2023년 1~2월 기획재정부 공동실사단',
        'KSA'  : '2023년 1~2월 기획재정부/닐슨아이큐코리아(유)'    
    }
    Q1_txt = {
        'KMAC' : '0점(전혀 그렇지 않다) 부터 1,2,3,4,5,6,7,8,9,10점(매우 그렇다)까지 동의하시는 정도가 클수록 높은 점수, 작을수록 낮은 점수를 주시면 됩니다.',
        'KSA'  : '0점(전혀 그렇지 않다), 5점(보통), 10점(매우 그렇다)까지 동의하시는 정도가 클수록 높은 점수, 작을수록 낮은 점수를 주시면 됩니다.'
    }

    Q5_txt = {
        'KMAC' : '0점(매우 나쁘다) 부터 1,2,3,4,5,6,7,8,9,10점(매우 좋다)까지 긍정적일수록 높은 점수, 부정적일수록 낮은 점수를 주시면 됩니다.',
        'KSA'  : '0점(매우 나쁘다), 5점(똑같다), 10점(매우 좋다)까지 긍정적일수록 높은 점수, 부정적 일수록 낮은 점수를 주시면 됩니다.'
    }

    Q6_txt = {
        'KMAC' : '0점(전혀 그렇지 않다) 부터 1,2,3,4,5,6,7,8,9,10점(매우 그렇다)까지 동의하시는 정도가 클수록 높은 점수, 작을수록 낮은 점수를 주시면 됩니다.',
        'KSA'  : '0점(전혀 그렇지 않다), 5점(보통이다), 10점(매우 그렇다)까지 동의하시는 정도가 클수록 높은 점수, 작을수록 낮은 점수를 주시면 됩니다.'
    }

    QQQ1_cell = new_ws.cell(16, 4)
    QQQ1_cell.value = QQQ1_cell.value.format(word=QQQ1_txt[division])

    Q1_cell = new_ws.cell(35, 4)
    Q1_cell.value = Q1_cell.value.format(word=Q1_txt[division])

    Q5_cell = new_ws.cell(46, 4)
    Q5_cell.value = Q5_cell.value.format(word=Q5_txt[division])

    Q6_cell = new_ws.cell(51, 4)
    Q6_cell.value = Q6_cell.value.format(word=Q6_txt[division])

    # KSA Q8 세팅
    if division == 'KSA' :
        max_array_length = max([len(i) for i in change_cells['Q8'].values()])

        # Q8 row setting
        Q8X1_qid_cell = new_ws.cell(58, 3)
        Q8X1_qid_cell.value = Q8X1_qid_cell.value%(max_array_length)

        # exmpample setting
        Q8X1_ex = [f'({i}) <span class=\'qnum\'>{i}</span>. <span class=\'row_text\'></span>' for i in range(1, max_array_length+1)]
        Q8X1_ex_cell = new_ws.cell(58, 6)
        Q8X1_ex_cell.value = Q8X1_ex_cell.value%('\n'.join(Q8X1_ex))

        Q8X2_ex = [f'({i}) {i}. <span class=\'SR_row_text\'></span>' for i in range(1, max_array_length+1)]
        Q8X2_ex_cell = new_ws.cell(59, 6)
        Q8X2_ex_cell.value = Q8X2_ex_cell.value%('\n'.join(Q8X2_ex))

        # js setting
        Q8_array = [f'if( QQQ14=={code} ){{ textArr = {arr};}}' for code, arr in change_cells['Q8'].items()]
        Q8X1_js_cell = new_ws.cell(57, 8)
        Q8X3_js_cell = new_ws.cell(60, 8)

        Q8X1_js_cell.value = Q8X1_js_cell.value%('\n'.join(Q8_array))
        Q8X3_js_cell.value = Q8X3_js_cell.value%('\n'.join(Q8_array))

    # 기관별 문항 타입 구분
    if division == 'KMAC' :
        # DQ2
        new_ws.cell(66, 7).value = None
        new_ws.delete_rows(68, 3)
        new_ws.delete_rows(56, 9)

    if division == 'KSA' :
        # DQ2X1, DQ2X2
        new_ws.delete_rows(67, 1)
        new_ws.delete_rows(55, 1)

    info_cell = new_ws.cell(33, 8)
    info_cell.value = info_cell.value%('\n'.join(info_texts))

    # save
    save_filename = f'PCSI_{division}_{survey_name}.xlsx'
    new_wb.save(os.path.join(os.getcwd(), save_folder, save_filename))


    print('💠 PCSI 스마트 서베이 확인 사항')
    print('   - xls로 다시 저장할 것')
    print('   - SQ/DQ 밑 설문지별 수정되는 변수 확인 필요')
    print('   - SQ1/SQ2도 설문지에 따라 다를 수 있음')
    print('   - DQ2 문항 : KMAC은 개인/법인 상관없이 DQ2에서 직업만 확인')
    print('   - DQ2 문항 : KSA는 개인의 경우 직업, 법인의 경우 직원수를 질문')
    print('   - 실사 담당자 전화번호 확인')
    print('   - 실사 시작전에 히든 변수 display_yn(n) 설정 해줄 것')
    print('   - 쿼터 세팅 확인')