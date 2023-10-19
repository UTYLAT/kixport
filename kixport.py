import sys, yaml, argparse, subprocess, pathlib, os, json, zipfile, PyPDF2, csv
from dataclasses import dataclass
from typing import List

# TODO
# - generate BOM + POS for JLCPCB
# - generate variant specific schematic and fab drawing
# - generate interactive html bom + variant specific version of it
# - generate different versions of Gerber (i.e with or without X2 extensions)
# - check that area fills are up to date
# - check that there are no unrouted nets on the board
# - generate panels
# - generate rendered images
# - put version number in file names
# - put version number inside generated PDFs and other files where it might be useful
# - make settings overridable for a specific board

@dataclass
class Board:
    '''Class to keep track of all settings for a board'''
    name: str
    kicad_pro: pathlib.Path
    kicad_sch: pathlib.Path
    kicad_pcb: pathlib.Path
    assembly_dir: pathlib.Path
    build_dir: pathlib.Path
    variants: List[dict]

    def from_yaml(board, settings):
        kicad_pro = pathlib.Path(board['kicad_pro'])
        return Board(
            name = board['name'],
            kicad_pro = kicad_pro,
            kicad_sch = kicad_pro.with_suffix('.kicad_sch'),
            kicad_pcb = kicad_pro.with_suffix('.kicad_pcb'),
            assembly_dir = pathlib.Path(f"{settings['assembly_dir']}/{board['name']}"),
            build_dir = pathlib.Path(f"{settings['build_dir']}/{board['name']}"),
            variants = board.get('variants')
        )

def read_version(board_pro):
    with open(board_pro, 'r') as file:
        js = json.load(file)
    return js['text_variables']['VERSION']

def mk_bom_xml(board_sch, bom_filename):
    print(f"--- Exporting XML netlist")
    subprocess.check_call(['kicad-cli', 'sch', 'export', 'python-bom', '-o', bom_filename, board_sch])

def mk_schematic_pdf(board_sch, pdf_filename):
    print(f"--- Generating schematic PDF")
    subprocess.check_call(['kicad-cli', 'sch', 'export', 'pdf', '-o', pdf_filename, board_sch])

def mk_gerber(board_pcb, gerber_dir):
    print(f"--- Generating gerber files")
    os.makedirs(gerber_dir, exist_ok=True)
    subprocess.check_call([
        'kicad-cli', 'pcb', 'export', 'gerbers', '-o', gerber_dir,
        '--use-drill-file-origin', board_pcb
    ])
    subprocess.check_call([
        'kicad-cli', 'pcb', 'export', 'drill', '-o', f"{gerber_dir}/",
        board_pcb
    ])

def mk_gerber_zip(gerber_dir, gerber_zip):
    print(f"--- Generating zip file for gerbers")
    with zipfile.ZipFile(gerber_zip, 'w') as zip:
        for file_path in gerber_dir.iterdir():
            zip.write(file_path, arcname=file_path.name)

def run_kibom(board_bom_xml: pathlib.Path, ini_filename: pathlib.Path, variant: str, bom_filename: pathlib.Path):
    print(f"--- Running KiBOM")
    bom_dir = bom_filename.parent.resolve()
    cmd = ['kibom', '--cfg', ini_filename, '-d', bom_dir]
    if variant:
        cmd += ['-r', variant]
    cmd += [board_bom_xml, bom_filename.name]
    print(f"KiBOM: {cmd}")
    subprocess.check_call(cmd)

def mk_fab_pdf(board: Board, settings, fab_filename: pathlib.Path):
    print(f"--- Generating fabrication PDF")
    writer = PyPDF2.PdfWriter()
    for fab in settings['outputs']['fab']:
        pdf_file = board.build_dir / f"{board.name}-fab-{fab['name']}.pdf"
        cmd = ['kicad-cli', 'pcb', 'export', 'pdf', '--output', pdf_file,
               '--layers', fab['layers'], board.kicad_pcb]
        if fab.get('mirror', False):
            cmd += ['--mirror']
        if fab['include_border_title']:
            cmd += ['--include-border-title']
        subprocess.check_call(cmd)
        reader = PyPDF2.PdfReader(pdf_file)
        writer.append(reader)
    with open(fab_filename, "wb") as file:
        writer.write(file)

def mk_step(board: Board, settings, step_path: pathlib.Path):
    print(f"--- Generating STEP file")
    subprocess.check_call(['kicad-cli', 'pcb', 'export', 'step', '--subst-models',
                           '--output', step_path, board.kicad_pcb])

def mk_pos(board: Board, settings, pos_path: pathlib.Path):
    print(f"--- Generating PnP position file (JLCPCB)")
    subprocess.check_call(['kicad-cli', 'pcb', 'export', 'pos', '--units', 'mm',
                           '--output', pos_path, board.kicad_pcb, '--format', pos_path.suffix[1:]])

# Read the generated CSV and change to match JLCPCB expectations
def mk_pos_jlcpcb(board: Board, settings, pos_path: pathlib.Path, jlcpcb_path: pathlib.Path):
    with open (pos_path, 'r') as input_file, open(jlcpcb_path, 'w') as output_file:
        input = csv.reader(input_file)
        output = csv.writer(output_file)

        # Skip first line with headers
        next(input, None)
        # Write the header JLCPCB expects
        output.writerow([u'Designator', u'Val', u'Package', u'Mid X', u'Mid y', u'Rotation', u'Layer'])
        # Copy the rest of the output
        for row in input:
            output.writerow(row)

def build_board(board: Board, settings):
    board_version = read_version(board.kicad_pro)
    print(f'Building {board.name} {board_version}')
    os.makedirs(board.assembly_dir, exist_ok=True)
    os.makedirs(board.build_dir, exist_ok=True)

    mk_schematic_pdf(board.kicad_sch, board.assembly_dir / f"{board.name}-schematic.pdf")

    gerber_dir = board.build_dir / f"{board.name}-gerber"
    mk_gerber(board.kicad_pcb, gerber_dir)

    gerber_zip = board.assembly_dir / f"{board.name}-gerber.zip"
    mk_gerber_zip(gerber_dir, gerber_zip)

    board_bom_xml = board.build_dir / f"{board.name}-bom.xml"
    mk_bom_xml(board.kicad_sch, board_bom_xml)

    for bom in settings['outputs']['kibom']:
        for format in bom['formats']:
            for variant in board.variants or [{'name': None, 'variant': None}]:
                if variant['variant'] != None:
                    bom_filename = f"{board.name}-{variant['variant']}-{bom['file_id']}.{format}"
                else:
                    bom_filename = f"{board.name}-{bom['file_id']}.{format}"
                run_kibom(board_bom_xml, bom['ini'], variant['variant'], board.assembly_dir / bom_filename)

    mk_fab_pdf(board, settings, board.assembly_dir / f"{board.name}-fab.pdf")

    mk_step(board, settings, board.assembly_dir / f"{board.name}.step")

    mk_pos(board, settings, board.build_dir / f"{board.name}-pos.csv")
    mk_pos_jlcpcb(board, settings, board.build_dir / f"{board.name}-pos.csv", board.assembly_dir / f"{board.name}-jlcpcb-cpl.csv")

def main(argv):
    with open(argv[1], 'r') as file:
        config = yaml.safe_load(file)

    settings = config['settings']

    for board in config['boards']:
        build_board(Board.from_yaml(board, settings), settings)

if __name__ == '__main__':
    sys.exit(main(sys.argv))
