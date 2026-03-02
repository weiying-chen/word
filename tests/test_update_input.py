from pathlib import Path

import update_input
import generate_news


def test_parse_source_txt_extracts_super_people_section(tmp_path: Path) -> None:
    source_txt = tmp_path / 'source.txt'
    source_txt.write_text(
        '\n'.join(
            [
                '建議標題：',
                'Sample title',
                'SUPER_PEOPLE：',
                '病患 | 羅伯托',
                'Roberto',
                'Patient',
                '',
                '醫師 | 林醫師',
                'Dr. Lin',
                'Physician',
                'Harbor Clinic',
                '',
                '字幕：',
                '1_0001',
                '中文內文。',
                'English body.',
            ]
        ),
        encoding='utf-8',
    )

    fields, body = update_input.parse_source_txt(source_txt)

    assert fields['TITLE_SUGGESTED'] == 'Sample title'
    assert fields['SUPER_PEOPLE'] == '\n'.join(
        [
            '病患 | 羅伯托',
            'Roberto',
            'Patient',
            '',
            '醫師 | 林醫師',
            'Dr. Lin',
            'Physician',
            'Harbor Clinic',
        ]
    )
    assert body == '1_0001\n中文內文。\nEnglish body.'


def test_write_input_emits_super_people_before_body(tmp_path: Path) -> None:
    output_path = tmp_path / 'input.txt'

    update_input.write_input(
        output_path,
        title='Sample title',
        url='https://example.com/news',
        summary='Summary text',
        time_range='( 11/16~17 )',
        fields={
            'SUPER_PEOPLE': '\n'.join(
                ['病患 | 羅伯托', 'Roberto', 'Patient']
            )
        },
        body='1_0001\nEnglish body.',
    )

    text = output_path.read_text(encoding='utf-8')

    assert 'SUPER_PEOPLE:\n病患 | 羅伯托\nRoberto\nPatient\n\nBODY:\n1_0001\nEnglish body.' in text


def test_generate_news_parse_input_preserves_super_people(tmp_path: Path) -> None:
    input_path = tmp_path / 'news_input.txt'
    input_path.write_text(
        '\n'.join(
            [
                'TITLE: Sample News Title',
                'SUMMARY:',
                'Summary line.',
                '',
                'SUPER_PEOPLE:',
                '病患 | 羅伯托',
                'Roberto',
                'Patient',
                '',
                'BODY:',
                '1_0001',
                '中文內文。',
                'English line.',
            ]
        ),
        encoding='utf-8',
    )

    data = generate_news.parse_input(input_path)

    assert data['SUPER_PEOPLE'] == '病患 | 羅伯托\nRoberto\nPatient'
    assert data['BODY'] == '1_0001\n中文內文。\nEnglish line.'
