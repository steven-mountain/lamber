"""Compare actual card/page DOCX outputs, resolving generated drawing relationship IDs."""
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'


def read(path):
    with ZipFile(path) as archive:
        xml = archive.read('word/document.xml')
        document = ET.fromstring(xml)
        relationships = {item.attrib['Id']: item.attrib['Target'] for item in ET.fromstring(archive.read('word/_rels/document.xml.rels'))}
        texts = [node.text or '' for node in document.iter(W + 't')]
        images = []
        for node in document.iter(A + 'blip'):
            target = relationships[node.attrib[R + 'embed']]
            member = target.lstrip('/') if target.startswith('/') else 'word/' + target
            digest = hashlib.sha256(archive.read(member)).hexdigest()
            images.append(digest)
            node.attrib[R + 'embed'] = digest
        for node in document.iter(WP + 'docPr'):
            node.attrib.pop('id', None)
            node.attrib.pop('name', None)
        return {'path': Path(path).name, 'fileSHA256': hashlib.sha256(Path(path).read_bytes()).hexdigest(),
                'bodySHA256': hashlib.sha256(xml).hexdigest(), 'texts': texts, 'images': images,
                'normalizedBody': ET.tostring(document)}


def compare(card, page):
    left, right = read(card), read(page)
    result = {'card': {key: left[key] for key in ('path', 'fileSHA256', 'bodySHA256')},
              'page': {key: right[key] for key in ('path', 'fileSHA256', 'bodySHA256')},
              'textNodeCount': len(left['texts']), 'textNodesIdentical': left['texts'] == right['texts'],
              'imagesInDocumentOrderIdentical': left['images'] == right['images'],
              'imageSHA256InDocumentOrder': left['images'],
              'normalizedBodyIdentical': left['normalizedBody'] == right['normalizedBody']}
    assert all(result[key] for key in ('textNodesIdentical', 'imagesInDocumentOrderIdentical', 'normalizedBodyIdentical')), result
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    result = {kind: compare(args.directory / f'card-{kind}.docx', args.directory / f'page-{kind}.docx') for kind in ('complete', 'missing')}
    encoded = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(encoded + '\n')
    print(encoded)
