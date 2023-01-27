#!/usr/bin/env python3

"""
Create csv file ready to upload to the drupal server, with all the
info from the excel with the list of substances.
"""

# Valid for the spreadsheet:
# ECDD1950_Substances considered (2023_01_23 TLR).xlsx
# (Others may have a different format and the program would need
# tweaking.

# To use with the "Substance record importer" feed at
# https://ecddrepository.org/en/admin/content/feed

import sys
import os
from collections import namedtuple
import openpyxl

from drugs_in_trs import drugs_in_trs

Record = namedtuple('Record', [
    'alternative_names', 'sessions', 'years', 'assessments', 'dclass',
    'effect', 'recom_ECDD', 'scheduling', 'link_report', 'link_questio',
    'trs_extract', 'trs', 'link_reviews'])



def main():
    if len(sys.argv) < 2:
        sys.exit('usage: %s <sheet file>' % sys.argv[0])

    sheet_fname = sys.argv[1]
    substances = get_substances(sheet_fname)

    write_csv('substances.csv', substances)


def write_csv(fname, substances):
    fout = open(fname, 'wt')

    fields = [
        'title',
        'field_drug_name',
        'field_alternative_names',
        'field_year',
        'field_year_s_and_type_of_review_',
        'field_drug_class',
        'field_drug_effect',
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
        path = 'https://ecddrepository.org/sites/default/files/'
        return (path + txt.lower()) if txt else ''

    trs_fname = {  # name of the file that corresponds to the TRS number
        '21': 'WHO_TRS_21.pdf',
        '57': 'WHO_TRS_57.pdf',
        '76': 'WHO_TRS_76.pdf',
        '95': 'WHO_TRS_95.pdf',
        '102': 'WHO_TRS_102.pdf',
        '116': 'WHO_TRS_116.pdf',
        '142': 'WHO_TRS_142.pdf',
        '160': 'WHO_TRS_160.pdf',
        '188': 'WHO_TRS_188.pdf',
        '211': 'WHO_TRS_211.pdf',
        '229': 'WHO_TRS_229.pdf',
        '273': 'WHO_TRS_273.pdf',
        '312': 'WHO_TRS_312.pdf',
        '343': 'WHO_TRS_343.pdf',
        '407': 'WHO_TRS_407.pdf',
        '437': 'WHO_TRS_437.pdf',
        '460': 'WHO_TRS_460.pdf',
        '526': 'WHO_TRS_526.pdf',
        '551': 'WHO_TRS_551.pdf',
        '729': 'WHO_TRS_729.pdf',
        '741': 'WHO_TRS_741.pdf',
        '761': 'WHO_TRS_761.pdf',
        '775': 'WHO_TRS_775.pdf',
        '787': 'WHO_TRS_787.pdf',
        '808': 'WHO_TRS_808.pdf',
        '836': 'WHO_TRS_836.pdf',
        '856': 'WHO_TRS_856.pdf',
        '873': 'WHO_TRS_873.pdf',
        '903': 'WHO_TRS_903.pdf',
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
        r.alternative_names,
        f'{r.years[-1]}-01-01',  # for "field_year" (it's of type date)
        build_sessions_years_assessments(r),  # for "field_year_s_and..."
        r.dclass,
        r.effect,
        r.recom_ECDD,
        r.scheduling,
        add_path(r.link_report),
        add_path(r.link_questio),
        r.trs_extract or get_recommendation(name, r.trs),
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
        return '"' + txt.replace('"', '""') + '"'
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
                print(f'Skipping merge - in row {i+1}, substance {name}: {e}')
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

    last_old = get_session_number(record_old.sessions[-1])
    last_new = get_session_number(record_new.sessions[-1])
    assert last_old < last_new, \
        f'not in chronological order: {last_old}, {last_new}'

    return record_new._replace(
        sessions=(record_old.sessions + record_new.sessions),
        years=(record_old.years + record_new.years),
        assessments=(record_old.assessments + record_new.assessments),
        link_reviews=(record_old.link_reviews + record_new.link_reviews))


def get_session_number(session):
    # '12th ECDD' -> 12
    assert (session.endswith(' ECDD') and
            session[-7:-5] in ['st', 'nd', 'rd', 'th']), \
        'session has incorrect ending: %r' % session
    return int(session[:-len('xx ECDD')])


def get_info(row):
    "Return the name of the substance and a record with all its info"
    # To extract the value at a given column.
    def get_index(cs):  # cs can be 'AC' for example
        i = ord(cs[-1]) - ord('A')
        unit = ord('Z') - ord('A') + 1
        for c in cs[-2::-1]:  # all the other in reverse
            i += unit * (ord(c) - ord('A') + 1)
            unit *= unit
        return i

    def value(c):
        assert c.isupper()
        v = row[get_index(c)].value

        if not v:
            return ''

        v = str(v).strip()
        replacements = [('‐', '-')]  # funny characters in the excel
        for r_from, r_to in replacements:
            v = v.replace(r_from, r_to)

        return v

    # To extract the link at a given column.
    def link(c):
        assert c.isupper()
        hl = row[get_index(c)].hyperlink
        return get_fname(hl.target) if hl else ''

    # Each letter contains the value at that column.
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W = \
        [value(c) for c in 'ABCDEFGHIJKLMNOPQRSTUVW']
    AF = value('AF')

    # Time to gather all that info!
    name = A

    trs = AF.split()[-1] if AF else ''  # TRS number (e.g. '942')

    record = Record(
        alternative_names=B,
        sessions=[G],
        years=[H],
        assessments=[J],
        dclass=F.lower(),
        effect=E,
        recom_ECDD=K,
        scheduling=L,
        link_report=link('M'),
        link_questio=link('N'),
        link_reviews=[link(c) for c in
                      ['X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE'] if link(c)],
        trs_extract=extract_sections([P, Q, R, S, T, U, V, W]),
        trs=trs)

    return name, record


def get_recommendation(name, trs):
    "Return the info for substance name that appears in that TRS"
    # Not really just the recommendation, but all of it.
    fname = f'extracted_from_trs/{trs}.txt'

    if not os.path.exists(fname):
        print(f'Not reading TRS from nonexistent file {fname}')
        return ''

    drugs = drugs_in_trs(fname)

    name = simplify_name(name)
    drugs = {simplify_name(k): v for k,v in drugs.items()}

    if name not in drugs:
        print(f'In TRS {trs}, missing {repr(name)} in {sorted(drugs)}')
        return ''

    return to_html(drugs[name].items())


def extract_sections(sections):
    s1, s2, s3, s4, s5, s6, s7, s8 = sections
    sections = [
        ('ECDD Technical summary', s1),
        ('Substance identification', s2),
        ('WHO review history', s3),
        ('Similarity to known substances and effects on the CNS', s4),
        ('Dependence potential', s5),
        ('Actual abuse and or/evidence of likelihood of abuse', s6),
        ('Therapeutic usefulness', s7),
        ('Recommendation', s8)]
    return to_html(sections)

def to_html(sections):
    p = lambda txt: '<br /><br />'.join(part for part in txt.split('\n'))
    return '<br /><br />'.join(f'<b><i>{name}</i></b><br />{p(text)}'
                   for name,text in sections if text)


def simplify_name(name):
    "Return a simplified version of name"
    # 'α-Lisdexamphetamine (INN) ' -> 'alpha-lisdexamfetamine'
    name = name.strip()  # just in case
    replacements = [('α', 'alpha'),
                    ('γ', 'gamma'),
                    ('pheta', 'feta')]
    for r_from, r_to in replacements:
        name = name.replace(r_from, r_to)
    if name.endswith(')'):
        return name[:name.rfind('(')].strip().lower()
    else:
        return name.lower()


def get_fname(link):
    # 'https://a/CriticalReview_5FPB22.pdf?ua=1' -> 'CriticalReview_5FPB22.pdf'
    qs_pos = link.rfind('?')
    link_noqs = link[:qs_pos] if qs_pos > 0 else link
    return link_noqs.split('/')[-1]



if __name__ == '__main__':
    main()
