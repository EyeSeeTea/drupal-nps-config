#!/usr/bin/env python3

"""
Create csv file ready to upload to the drupal server, with all the
info from the excel with the list of substances.
"""

import os
from collections import namedtuple
import openpyxl

from drugs_in_trs import drugs_in_trs

Record = namedtuple('Record', [
    'sessions', 'years', 'assessments', 'dclass', 'effect', 'use', 'recom_ECDD',
    'scheduling', 'link_report', 'link_questio', 'trs', 'link_reviews'])



def main():
    sheet_fname = ('../substances_list/'
                   'ECDD1950_Substances considered (2023_01_16 TLR).xlsx')
    substances = get_substances(sheet_fname)

    write_csv('substances.csv', substances)


def write_csv(fname, substances):
    fout = open(fname, 'wt')

    fields = [
        'title',
        'field_drug_name',
        'field_year',
        'field_year_s_and_type_of_review_',
        'field_drug_class',
        'field_drug_effect',
        'field_recognized_therapeutic_use',
        'field_ecdd_recommendation',
        'field_current_scheduling_status',
        'field_technical_information_most',
        'field_ms_questionnaire_report',
        'field_recommendation_from_trs_',
        'field_link_to_full_trs']

    fout.write(','.join(fields + [f'field_link{i}' for i in range(10)]) + '\n')

    for name, record in substances.items():
        fout.write(to_csv(name, record) + '\n')


def to_csv(name, record):
    "Return a csv line with all the information of the substance"

    def add_path(txt):
        #txt = txt.replace('--', '-')  # does drupal change this??
        return  ('sites/default/files/' + txt.lower()) if txt else ''

    trs_fname = {  # name of the file that corresponds to the TRS number
        '915': 'WHO_TRS_915.pdf',
        '942': 'WHO_TRS_942.pdf',
        '973': 'WHO_trs_973_eng.pdf',
        '991': 'WHO_TRS_991_eng.pdf',
        '998': 'WHO_TRS_998_eng.pdf',
        '1005': '9789241210140-eng.pdf',
        '1009': '9789241210188-eng.pdf',
        '1013': '9789241210225-eng.pdf',
        '1018': '9789241210270-eng.pdf',
        '1026': '9789240001848-eng.pdf',
        '1034': '9789240023024-eng.pdf',
        '1038': '9789240042834-eng.pdf'}

    r = record  # shortcut

    fields = [
        name,  # for "title"
        name,  # for "field_drug_name" (yes, repeated)
        f'{r.years[-1]}-01-01',  # for "field_year" (it's of type date)
        build_sessions_years_assessments(r),  # for "field_year_s_and..."
        class_code(r.dclass),
        r.effect,
        use_code(r.use),
        r.recom_ECDD,
        r.scheduling,
        add_path(r.link_report),
        add_path(r.link_questio),
        get_recommendation(name, r.trs),
        add_path(trs_fname.get(r.trs, ''))]

    links = [add_path(fname) for fname in r.link_reviews]

    return ','.join(escape_commas(field) for field in
                    (fields + links + ([''] * (10 - len(links)))))


def build_sessions_years_assessments(record):
    return ', '.join(f'{session} ({year}) - {assessment}'
                     for session,year,assessment in
                     zip(record.sessions, record.years, record.assessments))


def escape_commas(txt):
    if any(c in txt for c in ',“”"'):
        txt = txt.replace('“', '"')
        txt = txt.replace('”', '"')
        return '"' + txt.replace('"', r'\"') + '"'
    else:
        return txt


def get_substances(sheet_fname):
    "Return dict with records for each substance, read from the given sheet"
    substances = {}

    sheet = openpyxl.load_workbook(sheet_fname)['Full Sheet']

    for i,row in enumerate(sheet):
        if i == 0:
            continue  # first row is the header

        name, record = get_info(row)

        if name in substances:
            try:
                substances[name] = merge(substances[name], record)
            except AssertionError as e:
                print(f'Skipping merge - in row {i}, substance {name}: {e}')
        else:
            substances[name] = record

    return substances


def merge(record_old, record_new):
    "Merge, if possible, the two given records"
    # The stated drug class must match.
    assert record_old.dclass == record_new.dclass, \
        f'classes differ: {record_old.dclass}, {record_new.dclass}'

    # We don't do it, but could: require that the stated effects match.
    #assert record_old.effect == record_new.effect, \
    #    f'effects differ: {record_old.effect}, {record_new.effect}'

    last_old, last_new = record_old.sessions[-1], record_new.sessions[-1]
    assert last_old < last_new, \
        f'not in chronological order: {last_old}, {last_new}'

    return record_new._replace(
        sessions=(record_old.sessions + record_new.sessions),
        years=(record_old.years + record_new.years),
        assessments=(record_old.assessments + record_new.assessments),
        link_reviews=(record_old.link_reviews + record_new.link_reviews))


def get_info(row):
    "Return the name of the substance and a record with all its info"
    # To extract the value at a given column.
    def value(c):
        v = row[ord(c.upper()) - ord('A')].value

        if not v:
            return ''

        v = str(v).strip()
        replacements = [('‐', '-')]  # funny characters in the excel
        for r_from, r_to in replacements:
            v = v.replace(r_from, r_to)

        return v

    # To extract the link at a given column.
    def link(c):
        hl = row[ord(c.upper()) - ord('A')].hyperlink
        return get_fname(hl.target) if hl else ''

    # Each letter contains the value at that column.
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V = \
        [value(c) for c in 'ABCDEFGHIJKLMNOPQRSTUV']

    # Time to gather all that info!
    name = C

    trs = V.split()[-1] if V else ''  # TRS number (e.g. '942')

    if G == '':
        print('Found a row without therapeutic use specified.')

    record = Record(
        sessions=[A],
        years=[B],
        assessments=[H],
        dclass=F,
        effect=E,
        use=G or 'None',
        recom_ECDD=I,
        scheduling=K,
        link_report=link('L'),
        link_questio=link('T'),
        link_reviews=[link(c) for c in 'MNOPQRS' if link(c)],
        trs=trs)

    return name, record


def get_recommendation(name, trs):
    "Return the info for substance name that appears in that TRS"
    # Not really just the recommendation, but all of it.
    fname = f'../pdfs/links_to_trs/{trs}.txt'

    if not os.path.exists(fname):
        print(f'Not reading TRS from nonexistent file {fname}')
        return ''

    drugs = drugs_in_trs(fname)

    name = simplify_name(name)
    drugs = {simplify_name(k): v for k,v in drugs.items()}

    if name not in drugs:
        print(f'In TRS {trs}, missing {repr(name)} in {sorted(drugs)}')
        return ''

    p = lambda txt: ''.join(f'{part}<br /><br />' for part in txt.split('\n'))
    return ''.join(f'<b><i>{k}</i></b><br />{p(v)}'
                   for k,v in drugs[name].items())
    # If we only wanted the recommendation part, it would be:
    #return drugs[name].get('Recommendation', '')


def simplify_name(name):
    "Return a simplified version of name"
    # 'α-Lisdexamphetamine (INN) ' -> 'alpha-lisdexamfetamine'
    name = name.strip()  # just in case
    replacements = [('α', 'alpha'),
                    ('γ', 'gamma'),
                    ('pheta', 'feta')]
    for r_from, r_to in replacements:
        name = name.replace(r_from, r_to)
    if name.endswith('('):
        return name[:name.rfind('(')].strip().lower()
    else:
        return name.lower()


def use_code(txt):
    code_start = 10  #82  # it happens to be the one in drupal for the 1st one
    uses = ['therapeutic use', 'none']
    return str(code_start + uses.index(txt.lower().strip()))


def class_code(txt):
    code_start = 1 #84  # it happens to be the one in drupal for the 1st one
    classes = ['benzodiazepine', 'cannabinoid', 'dissociative', 'hallucinogen',
               'insufficient information', 'opioids', 'other', 'sedatives',
               'stimulant']
    return str(code_start + classes.index(txt.lower().strip()))


def get_fname(link):
    # 'https://a/CriticalReview_5FPB22.pdf?ua=1' -> 'CriticalReview_5FPB22.pdf'
    qs_pos = link.rfind('?')
    link_noqs = link[:qs_pos] if qs_pos > 0 else link
    return link_noqs.split('/')[-1]



if __name__ == '__main__':
    main()
