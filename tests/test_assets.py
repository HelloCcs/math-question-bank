from pathlib import Path
from PIL import Image
from question_bank.assets import build_asset_manifest, image_kind, safe_asset_path

def test_safe_asset_path_rejects_traversal(tmp_path):
    assert safe_asset_path("previews/a/image.png", tmp_path) == (tmp_path/"previews/a/image.png").resolve()
    assert safe_asset_path("../../secret.txt", tmp_path) is None

def test_image_kind_and_manifest(tmp_path):
    inline=tmp_path/"inline.png"; block=tmp_path/"block.png"
    Image.new("RGB",(300,60),"white").save(inline); Image.new("RGB",(600,400),"white").save(block)
    assert image_kind(inline)=="inline-formula"; assert image_kind(block)=="block-figure"
    manifest=build_asset_manifest('<p><img class="inline-math inline-formula" src="a.png"></p><img class="block-figure" src="b.png">')
    assert manifest==[{"src":"a.png","kind":"inline-formula"},{"src":"b.png","kind":"block-figure"}]
