import random
import difflib
import genanki
import openpyxl
import openpyxl.cell.read_only
import openpyxl.cell.rich_text
from os import PathLike
from typing import Union, Dict, IO, BinaryIO

my_model = genanki.Model(
    1607392319,
    'Simple Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'},
    ],
    templates=[
        {
            'name': 'Card 1',
            'qfmt': '{{Question}}',
            'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}',
        },
    ])


def is_empty_cell(
    row: Dict[str, openpyxl.cell.read_only.ReadOnlyCell],
    col_name: str,
) -> bool:
    return col_name not in row or row[col_name].value is None


def convert_xlsx_to_apkg(
    input_file: Union[PathLike, str, BinaryIO, IO[bytes]],
    output_file: Union[PathLike, str, BinaryIO, IO[bytes]],
    deck_name: str
):
    deck_id = random.randrange(1 << 30, 1 << 31)
    my_deck = genanki.Deck(deck_id, deck_name)

    wb = openpyxl.load_workbook(input_file, read_only=True, rich_text=True)
    sheet = wb.active

    # Read table's rows
    header = None
    for i, row in enumerate(sheet):
        if i == 0:
            # Read table's header
            header = [str(c.value) if c.value is not None else None for c in row]
            assert 'Source' in header, '`Source` column is missing'
            assert 'Target1' in header, '`Target1` column is missing. At least one Target column must be present'
        else:
            row = dict(zip(header, row))

            # Get source phrase, add accent mark(s)
            if is_empty_cell(row, 'Source'):
                src = ''
            elif isinstance(row['Source'].value, str):
                src = row['Source'].value
            elif isinstance(row['Source'].value, openpyxl.cell.rich_text.CellRichText):
                # Add accent mark for each underlined character
                src = ''.join(
                    (
                        ''.join(c + u'\u0301' for c in p.text)  # add accent mark
                        if isinstance(p, openpyxl.cell.rich_text.TextBlock) and p.font.u  # if underlined
                        else str(p)  # otherwise, just convert to string
                    ) for p in row['Source'].value
                )
            else:
                raise ValueError(f'Unexpected type {type(row["Source"].value)} for Source cell in row {i}')

            # Get example usage
            if is_empty_cell(row, 'Example'):
                example = ''
            else:
                example = str(row['Example'].value)
                example_words = example.split()
                if example_words:
                    # Highlight the word with the highest similarity to the source word
                    example_sims = [difflib.SequenceMatcher(None, src, word).ratio() for word in example_words]
                    max_example_sim = max(example_sims)
                    example = ' '.join(
                        f'<b>{word}</b>' if sim == max_example_sim else word
                        for word, sim in zip(example_words, example_sims)
                    )

            # Get target phrase(s)
            targets = []
            i = 1
            while f'Target{i}' in header:
                if not is_empty_cell(row, f'Target{i}'):
                    target = str(row[f'Target{i}'].value)
                    targets.append(target)
                i += 1

            # Generate a flash card
            question = f'{src}<br><br><i>{example}</i>' if example else src
            answer = ' | '.join(targets)
            if question and answer:
                my_note = genanki.Note(
                    model=my_model,
                    fields=[question, answer]
                )
                my_deck.add_note(my_note)

    # Export flash cards
    genanki.Package(my_deck).write_to_file(output_file)
