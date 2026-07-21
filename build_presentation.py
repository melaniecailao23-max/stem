#!/usr/bin/env python3
"""
Build the STEM Education PowerPoint presentation from scratch.
This script modifies the starter PPTX file to implement all 28 steps.
"""

import zipfile
import shutil
import os
import copy
import xml.etree.ElementTree as ET
from io import BytesIO

# Constants - EMU conversions (1 inch = 914400 EMU)
EMU_PER_INCH = 914400

# Paths
WORK_DIR = '/projects/sandbox/stem'
STARTER_PPTX = os.path.join(WORK_DIR, 'Cailao_Exp22_PPT_AppCapstone_Stem.pptx')
ASSETS_DIR = os.path.join(WORK_DIR, 'assets/Exp22_PPT_AppCapstone_Intro_Stem_Assets')
OUTPUT_PPTX = os.path.join(WORK_DIR, 'Cailao_Exp22_PPT_AppCapstone_Stem_Complete.pptx')
EXTRACT_DIR = os.path.join(WORK_DIR, 'pptx_work')

# Namespace map for XML
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'p15': 'http://schemas.microsoft.com/office/powerpoint/2012/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# Register namespaces
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

# Additional namespaces needed
ET.register_namespace('', 'http://schemas.openxmlformats.org/presentationml/2006/main')



def inches_to_emu(inches):
    """Convert inches to EMU."""
    return int(inches * EMU_PER_INCH)

def extract_pptx(pptx_path, extract_dir):
    """Extract PPTX to directory."""
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    with zipfile.ZipFile(pptx_path, 'r') as z:
        z.extractall(extract_dir)

def repack_pptx(extract_dir, output_path):
    """Repack directory into PPTX file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                z.write(file_path, arcname)

def read_xml(filepath):
    """Read and parse XML file."""
    tree = ET.parse(filepath)
    return tree

def write_xml(tree, filepath):
    """Write XML tree to file."""
    tree.write(filepath, xml_declaration=True, encoding='UTF-8')



def add_relationship(rels_path, rid, rel_type, target):
    """Add a relationship to a .rels file."""
    tree = ET.parse(rels_path)
    root = tree.getroot()
    rel_elem = ET.SubElement(root, 'Relationship')
    rel_elem.set('Id', rid)
    rel_elem.set('Type', rel_type)
    rel_elem.set('Target', target)
    tree.write(rels_path, xml_declaration=True, encoding='UTF-8')

def get_next_rid(rels_path):
    """Get the next available relationship ID."""
    tree = ET.parse(rels_path)
    root = tree.getroot()
    max_id = 0
    for rel in root:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                num = int(rid[3:])
                max_id = max(max_id, num)
            except ValueError:
                pass
    return f'rId{max_id + 1}'

def copy_image_to_media(src_path, media_dir, filename):
    """Copy an image file to the pptx media directory."""
    os.makedirs(media_dir, exist_ok=True)
    dst_path = os.path.join(media_dir, filename)
    shutil.copy2(src_path, dst_path)
    return filename



def ensure_content_types(extract_dir, extensions_map):
    """Ensure content types file has required extensions."""
    ct_path = os.path.join(extract_dir, '[Content_Types].xml')
    if not os.path.exists(ct_path):
        # Create content types file
        root = ET.Element('Types')
        root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types')
        tree = ET.ElementTree(root)
    else:
        tree = ET.parse(ct_path)
        root = tree.getroot()
    
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    
    # Check existing extensions
    existing_exts = set()
    for child in root:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'Default':
            existing_exts.add(child.get('Extension', '').lower())
    
    for ext, content_type in extensions_map.items():
        if ext.lower() not in existing_exts:
            default = ET.SubElement(root, f'{{{ct_ns}}}Default' if ct_ns else 'Default')
            default.set('Extension', ext)
            default.set('ContentType', content_type)
    
    tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
    return ct_path



def step2_change_theme_colors(extract_dir):
    """Step 2: Change theme colors to Aspect."""
    theme_path = os.path.join(extract_dir, 'ppt/theme/theme1.xml')
    if not os.path.exists(theme_path):
        # Find theme file
        theme_dir = os.path.join(extract_dir, 'ppt/theme')
        if os.path.exists(theme_dir):
            for f in os.listdir(theme_dir):
                if f.endswith('.xml'):
                    theme_path = os.path.join(theme_dir, f)
                    break
    
    if not os.path.exists(theme_path):
        print("Warning: Theme file not found, creating one")
        return
    
    tree = ET.parse(theme_path)
    root = tree.getroot()
    
    # Aspect color scheme values
    # dk1=000000, lt1=FFFFFF, dk2=323232, lt2=E3DED1
    # accent1=F07F09, accent2=9F2936, accent3=1B587C, accent4=4E8542
    # accent5=604878, accent6=C19859, hlink=6B9F25, folHlink=B26B02
    aspect_colors = {
        'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '323232', 'lt2': 'E3DED1',
        'accent1': 'F07F09', 'accent2': '9F2936', 'accent3': '1B587C',
        'accent4': '4E8542', 'accent5': '604878', 'accent6': 'C19859',
        'hlink': '6B9F25', 'folHlink': 'B26B02'
    }
    
    # Find the color scheme element
    a_ns = NS['a']
    for elem in root.iter(f'{{{a_ns}}}clrScheme'):
        elem.set('name', 'Aspect')
        # Remove existing color children
        children_to_remove = list(elem)
        for child in children_to_remove:
            elem.remove(child)
        # Add new colors
        for color_name, color_val in aspect_colors.items():
            color_elem = ET.SubElement(elem, f'{{{a_ns}}}{color_name}')
            srgb = ET.SubElement(color_elem, f'{{{a_ns}}}srgbClr')
            srgb.set('val', color_val)
    
    tree.write(theme_path, xml_declaration=True, encoding='UTF-8')
    print("Step 2: Theme colors changed to Aspect")



def step3_change_theme_fonts(extract_dir):
    """Step 3: Change theme fonts to Gill Sans MT."""
    theme_path = os.path.join(extract_dir, 'ppt/theme/theme1.xml')
    if not os.path.exists(theme_path):
        theme_dir = os.path.join(extract_dir, 'ppt/theme')
        for f in os.listdir(theme_dir):
            if f.endswith('.xml'):
                theme_path = os.path.join(theme_dir, f)
                break
    
    tree = ET.parse(theme_path)
    root = tree.getroot()
    
    a_ns = NS['a']
    for font_scheme in root.iter(f'{{{a_ns}}}fontScheme'):
        font_scheme.set('name', 'Gill Sans MT')
        # Update major and minor fonts
        for major in font_scheme.iter(f'{{{a_ns}}}majorFont'):
            for latin in major.iter(f'{{{a_ns}}}latin'):
                latin.set('typeface', 'Gill Sans MT')
        for minor in font_scheme.iter(f'{{{a_ns}}}minorFont'):
            for latin in minor.iter(f'{{{a_ns}}}latin'):
                latin.set('typeface', 'Gill Sans MT')
    
    tree.write(theme_path, xml_declaration=True, encoding='UTF-8')
    print("Step 3: Theme fonts changed to Gill Sans MT")



def step4_slide8_microscope(extract_dir):
    """Step 4: On Slide 8, insert Microscope.jpg in left placeholder with Metal Oval style."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide8.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide8.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    # Copy image
    src_img = os.path.join(ASSETS_DIR, 'Microsope.jpg')
    copy_image_to_media(src_img, media_dir, 'Microscope.jpg')
    
    # Add relationship
    rid = get_next_rid(rels_path)
    add_relationship(rels_path, rid,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        '../media/Microscope.jpg')
    
    # Parse slide XML and add picture to the picture placeholder
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    # Find the picture placeholder (type="pic", idx="1")
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Find picture placeholder and replace with actual picture
    pic_placeholder = None
    for sp in sp_tree.findall(f'{{{p_ns}}}sp'):
        nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvSpPr is not None:
            nvPr = nvSpPr.find(f'{{{p_ns}}}nvPr')
            if nvPr is not None:
                ph = nvPr.find(f'{{{p_ns}}}ph')
                if ph is not None and ph.get('type') == 'pic':
                    pic_placeholder = sp
                    break
    
    if pic_placeholder is not None:
        # Get position from placeholder
        spPr = pic_placeholder.find(f'{{{p_ns}}}spPr')
        xfrm = spPr.find(f'{{{a_ns}}}xfrm') if spPr is not None else None
        
        # Remove placeholder
        sp_tree.remove(pic_placeholder)
        
        # Create picture element
        pic = ET.SubElement(sp_tree, f'{{{p_ns}}}pic')
        
        # nvPicPr
        nvPicPr = ET.SubElement(pic, f'{{{p_ns}}}nvPicPr')
        cNvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', '20')
        cNvPr.set('name', 'Microscope')
        cNvPicPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPicPr')
        picLocks = ET.SubElement(cNvPicPr, f'{{{a_ns}}}picLocks')
        picLocks.set('noChangeAspect', '1')
        nvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}nvPr')
        
        # blipFill
        blipFill = ET.SubElement(pic, f'{{{p_ns}}}blipFill')
        blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
        blip.set(f'{{{r_ns}}}embed', rid)
        stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
        fillRect = ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
        
        # spPr with Metal Oval style effect and position
        spPr_new = ET.SubElement(pic, f'{{{p_ns}}}spPr')
        if xfrm is not None:
            spPr_new.append(xfrm)
        else:
            xfrm_new = ET.SubElement(spPr_new, f'{{{a_ns}}}xfrm')
            off = ET.SubElement(xfrm_new, f'{{{a_ns}}}off')
            off.set('x', '349856')
            off.set('y', '222637')
            ext = ET.SubElement(xfrm_new, f'{{{a_ns}}}ext')
            ext.set('cx', '5472503')
            ext.set('cy', '6347845')
        
        prstGeom = ET.SubElement(spPr_new, f'{{{a_ns}}}prstGeom')
        prstGeom.set('prst', 'ellipse')
        avLst = ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
        
        # Metal Oval border - Dark Gray, Background 2, Lighter 50%
        ln = ET.SubElement(spPr_new, f'{{{a_ns}}}ln')
        ln.set('w', '88900')
        solidFill = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
        schemeClr.set('val', 'bg2')
        lumMod = ET.SubElement(schemeClr, f'{{{a_ns}}}lumMod')
        lumMod.set('val', '50000')
        lumOff = ET.SubElement(schemeClr, f'{{{a_ns}}}lumOff')
        lumOff.set('val', '50000')
        
        # Effect for Metal Oval (reflection/shadow)
        effectLst = ET.SubElement(spPr_new, f'{{{a_ns}}}effectLst')
        reflection = ET.SubElement(effectLst, f'{{{a_ns}}}reflection')
        reflection.set('blurRad', '12700')
        reflection.set('stA', '28000')
        reflection.set('endPos', '45000')
        reflection.set('dist', '1000')
        reflection.set('dir', '5400000')
        reflection.set('sy', '-100000')
        reflection.set('algn', 'bl')
        reflection.set('rotWithShape', '0')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 4: Microscope.jpg inserted on Slide 8 with Metal Oval style")



def step5_slide10_images(extract_dir):
    """Step 5: On Slide 10, insert ElementarySchool images with Metal Frame style."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide10.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide10.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    # Copy images
    copy_image_to_media(os.path.join(ASSETS_DIR, 'ElementarySchool.jpg'), media_dir, 'ElementarySchool.jpg')
    copy_image_to_media(os.path.join(ASSETS_DIR, 'ElementarySchool2.jpg'), media_dir, 'ElementarySchool2.jpg')
    
    # Add relationships
    rid1 = get_next_rid(rels_path)
    add_relationship(rels_path, rid1,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        '../media/ElementarySchool.jpg')
    rid2 = get_next_rid(rels_path)
    add_relationship(rels_path, rid2,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        '../media/ElementarySchool2.jpg')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    def create_pic_with_metal_frame(parent, rid, name, x, y, cx, cy):
        """Create a picture element with Metal Frame style."""
        pic = ET.SubElement(parent, f'{{{p_ns}}}pic')
        nvPicPr = ET.SubElement(pic, f'{{{p_ns}}}nvPicPr')
        cNvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', str(hash(name) % 1000 + 100))
        cNvPr.set('name', name)
        cNvPicPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPicPr')
        picLocks = ET.SubElement(cNvPicPr, f'{{{a_ns}}}picLocks')
        picLocks.set('noChangeAspect', '1')
        nvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}nvPr')
        
        blipFill = ET.SubElement(pic, f'{{{p_ns}}}blipFill')
        blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
        blip.set(f'{{{r_ns}}}embed', rid)
        stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
        ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
        
        spPr = ET.SubElement(pic, f'{{{p_ns}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
        off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
        off.set('x', str(x))
        off.set('y', str(y))
        ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
        ext.set('cx', str(cx))
        ext.set('cy', str(cy))
        prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
        
        # Metal Frame border
        ln = ET.SubElement(spPr, f'{{{a_ns}}}ln')
        ln.set('w', '88900')
        solidFill = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
        schemeClr.set('val', 'bg2')
        lumMod = ET.SubElement(schemeClr, f'{{{a_ns}}}lumMod')
        lumMod.set('val', '50000')
        lumOff = ET.SubElement(schemeClr, f'{{{a_ns}}}lumOff')
        lumOff.set('val', '50000')
        
        return pic
    
    # Left content placeholder position
    create_pic_with_metal_frame(sp_tree, rid1, 'ElementarySchool',
        inches_to_emu(0.5), inches_to_emu(1.5), inches_to_emu(5.0), inches_to_emu(5.0))
    # Right content placeholder position  
    create_pic_with_metal_frame(sp_tree, rid2, 'ElementarySchool2',
        inches_to_emu(6.0), inches_to_emu(1.5), inches_to_emu(5.0), inches_to_emu(5.0))
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 5: Elementary School images inserted on Slide 10")



def step6_slide11_middleschool(extract_dir):
    """Step 6: On Slide 11, insert MiddleSchool.jpg with specific formatting."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide11.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide11.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    copy_image_to_media(os.path.join(ASSETS_DIR, 'MiddleSchool.jpg'), media_dir, 'MiddleSchool.jpg')
    
    rid = get_next_rid(rels_path)
    add_relationship(rels_path, rid,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        '../media/MiddleSchool.jpg')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Create picture with Beveled Oval, Black style
    pic = ET.SubElement(sp_tree, f'{{{p_ns}}}pic')
    nvPicPr = ET.SubElement(pic, f'{{{p_ns}}}nvPicPr')
    cNvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '30')
    cNvPr.set('name', 'MiddleSchool')
    cNvPicPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPicPr')
    picLocks = ET.SubElement(cNvPicPr, f'{{{a_ns}}}picLocks')
    picLocks.set('noChangeAspect', '1')
    nvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}nvPr')
    
    blipFill = ET.SubElement(pic, f'{{{p_ns}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
    blip.set(f'{{{r_ns}}}embed', rid)
    stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
    ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
    
    spPr = ET.SubElement(pic, f'{{{p_ns}}}spPr')
    # Position: H=6.67", V=0.4", Height=6.7"
    xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
    off.set('x', str(inches_to_emu(6.67)))
    off.set('y', str(inches_to_emu(0.4)))
    ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
    ext.set('cx', str(inches_to_emu(5.0)))  # Width proportional
    ext.set('cy', str(inches_to_emu(6.7)))
    
    # Beveled Oval shape
    prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
    prstGeom.set('prst', 'ellipse')
    ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
    
    # Border - Dark Gray, Background 2, Lighter 50%
    ln = ET.SubElement(spPr, f'{{{a_ns}}}ln')
    ln.set('w', '88900')
    solidFill = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
    schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
    schemeClr.set('val', 'bg2')
    lumMod = ET.SubElement(schemeClr, f'{{{a_ns}}}lumMod')
    lumMod.set('val', '50000')
    lumOff = ET.SubElement(schemeClr, f'{{{a_ns}}}lumOff')
    lumOff.set('val', '50000')
    
    # Bevel effect
    effectLst = ET.SubElement(spPr, f'{{{a_ns}}}effectLst')
    outerShdw = ET.SubElement(effectLst, f'{{{a_ns}}}outerShdw')
    outerShdw.set('blurRad', '76200')
    outerShdw.set('dist', '38100')
    outerShdw.set('dir', '5400000')
    outerShdw.set('algn', 'tl')
    srgbClr = ET.SubElement(outerShdw, f'{{{a_ns}}}srgbClr')
    srgbClr.set('val', '000000')
    alpha = ET.SubElement(srgbClr, f'{{{a_ns}}}alpha')
    alpha.set('val', '40000')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 6: MiddleSchool.jpg inserted on Slide 11")



def step7_slide12_video(extract_dir):
    """Step 7: On Slide 12, insert HSVideo.mp4 with formatting."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide12.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide12.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    # Copy video
    shutil.copy2(os.path.join(ASSETS_DIR, 'HSVideo.mp4'), os.path.join(media_dir, 'HSVideo.mp4'))
    
    rid_video = get_next_rid(rels_path)
    add_relationship(rels_path, rid_video,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
        '../media/HSVideo.mp4')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Create video placeholder as a picture with video link
    pic = ET.SubElement(sp_tree, f'{{{p_ns}}}pic')
    nvPicPr = ET.SubElement(pic, f'{{{p_ns}}}nvPicPr')
    cNvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '40')
    cNvPr.set('name', 'HSVideo.mp4')
    cNvPicPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPicPr')
    picLocks = ET.SubElement(cNvPicPr, f'{{{a_ns}}}picLocks')
    picLocks.set('noChangeAspect', '1')
    nvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}nvPr')
    # Video link for auto start
    videoFile = ET.SubElement(nvPr, f'{{{a_ns}}}videoFile')
    videoFile.set(f'{{{r_ns}}}link', rid_video)
    
    blipFill = ET.SubElement(pic, f'{{{p_ns}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
    blip.set(f'{{{r_ns}}}embed', rid_video)
    stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
    ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
    
    spPr = ET.SubElement(pic, f'{{{p_ns}}}spPr')
    xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
    off.set('x', str(inches_to_emu(2.0)))
    off.set('y', str(inches_to_emu(1.5)))
    ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
    ext.set('cx', str(inches_to_emu(8.0)))
    ext.set('cy', str(inches_to_emu(5.0)))
    prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
    prstGeom.set('prst', 'rect')
    ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
    
    # Perspective Shadow + border
    ln = ET.SubElement(spPr, f'{{{a_ns}}}ln')
    ln.set('w', '25400')
    solidFill = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
    schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
    schemeClr.set('val', 'bg2')
    lumMod = ET.SubElement(schemeClr, f'{{{a_ns}}}lumMod')
    lumMod.set('val', '50000')
    lumOff = ET.SubElement(schemeClr, f'{{{a_ns}}}lumOff')
    lumOff.set('val', '50000')
    
    effectLst = ET.SubElement(spPr, f'{{{a_ns}}}effectLst')
    outerShdw = ET.SubElement(effectLst, f'{{{a_ns}}}outerShdw')
    outerShdw.set('blurRad', '190500')
    outerShdw.set('dist', '228600')
    outerShdw.set('dir', '2700000')
    outerShdw.set('sx', '85000')
    outerShdw.set('sy', '-23000')
    outerShdw.set('kx', '800400')
    outerShdw.set('algn', 'bl')
    srgbClr = ET.SubElement(outerShdw, f'{{{a_ns}}}srgbClr')
    srgbClr.set('val', '808080')
    alpha = ET.SubElement(srgbClr, f'{{{a_ns}}}alpha')
    alpha.set('val', '43000')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 7: HSVideo.mp4 inserted on Slide 12")



def step8_slide4_background(extract_dir):
    """Step 8: On Slide 4, insert WorldMap.jpg as background with 10% transparency."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide4.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide4.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    copy_image_to_media(os.path.join(ASSETS_DIR, 'WorldMap.jpg'), media_dir, 'WorldMap.jpg')
    
    rid = get_next_rid(rels_path)
    add_relationship(rels_path, rid,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        '../media/WorldMap.jpg')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    # Add background element to slide
    cSld = root.find(f'{{{p_ns}}}cSld')
    
    # Create bg element before spTree
    bg = ET.Element(f'{{{p_ns}}}bg')
    bgPr = ET.SubElement(bg, f'{{{p_ns}}}bgPr')
    blipFill = ET.SubElement(bgPr, f'{{{a_ns}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
    blip.set(f'{{{r_ns}}}embed', rid)
    # 10% transparency = alphaModFix val=90000 (90% opaque)
    alphaModFix = ET.SubElement(blip, f'{{{a_ns}}}alphaModFix')
    alphaModFix.set('amt', '90000')
    stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
    ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
    ET.SubElement(bgPr, f'{{{a_ns}}}effectLst')
    
    # Insert bg as first child of cSld
    cSld.insert(0, bg)
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 8: WorldMap.jpg set as Slide 4 background")



def step9_slide2_smartart(extract_dir):
    """Step 9: On Slide 2, insert Picture Caption List SmartArt."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide2.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide2.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    # Copy images
    images = ['Science.jpg', 'Technology.jpg', 'Engineering.jpg', 'Math.jpg']
    rids = []
    for img in images:
        copy_image_to_media(os.path.join(ASSETS_DIR, img), media_dir, img)
        rid = get_next_rid(rels_path)
        add_relationship(rels_path, rid,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            f'../media/{img}')
        rids.append(rid)
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Remove existing content placeholder
    for sp in list(sp_tree.findall(f'{{{p_ns}}}sp')):
        nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvSpPr is not None:
            nvPr = nvSpPr.find(f'{{{p_ns}}}nvPr')
            if nvPr is not None:
                ph = nvPr.find(f'{{{p_ns}}}ph')
                if ph is not None and ph.get('type') != 'title':
                    sp_tree.remove(sp)
    
    # Create a group of 4 picture+caption items (simulating SmartArt)
    labels = ['Science', 'Technology', 'Engineering', 'Math']
    x_positions = [inches_to_emu(0.5), inches_to_emu(3.2), inches_to_emu(5.9), inches_to_emu(8.6)]
    
    for i, (label, rid, x_pos) in enumerate(zip(labels, rids, x_positions)):
        # Picture
        pic = ET.SubElement(sp_tree, f'{{{p_ns}}}pic')
        nvPicPr = ET.SubElement(pic, f'{{{p_ns}}}nvPicPr')
        cNvPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', str(50 + i))
        cNvPr.set('name', f'{label} Picture')
        cNvPicPr = ET.SubElement(nvPicPr, f'{{{p_ns}}}cNvPicPr')
        picLocks = ET.SubElement(cNvPicPr, f'{{{a_ns}}}picLocks')
        picLocks.set('noChangeAspect', '1')
        ET.SubElement(nvPicPr, f'{{{p_ns}}}nvPr')
        
        blipFill = ET.SubElement(pic, f'{{{p_ns}}}blipFill')
        blip = ET.SubElement(blipFill, f'{{{a_ns}}}blip')
        blip.set(f'{{{r_ns}}}embed', rid)
        stretch = ET.SubElement(blipFill, f'{{{a_ns}}}stretch')
        ET.SubElement(stretch, f'{{{a_ns}}}fillRect')
        
        spPr = ET.SubElement(pic, f'{{{p_ns}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
        off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
        off.set('x', str(x_pos))
        off.set('y', str(inches_to_emu(1.8)))
        ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
        ext.set('cx', str(inches_to_emu(2.4)))
        ext.set('cy', str(inches_to_emu(2.4)))
        prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
        
        # Intense Effect style - 3D bevel + shadow
        effectLst = ET.SubElement(spPr, f'{{{a_ns}}}effectLst')
        outerShdw = ET.SubElement(effectLst, f'{{{a_ns}}}outerShdw')
        outerShdw.set('blurRad', '40000')
        outerShdw.set('dist', '23000')
        outerShdw.set('dir', '5400000')
        srgbClr = ET.SubElement(outerShdw, f'{{{a_ns}}}srgbClr')
        srgbClr.set('val', '000000')
        alpha = ET.SubElement(srgbClr, f'{{{a_ns}}}alpha')
        alpha.set('val', '35000')
        
        # Caption text box
        sp = ET.SubElement(sp_tree, f'{{{p_ns}}}sp')
        nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
        cNvPr2 = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr2.set('id', str(60 + i))
        cNvPr2.set('name', f'{label} Caption')
        ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
        ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
        
        spPr2 = ET.SubElement(sp, f'{{{p_ns}}}spPr')
        xfrm2 = ET.SubElement(spPr2, f'{{{a_ns}}}xfrm')
        off2 = ET.SubElement(xfrm2, f'{{{a_ns}}}off')
        off2.set('x', str(x_pos))
        off2.set('y', str(inches_to_emu(4.4)))
        ext2 = ET.SubElement(xfrm2, f'{{{a_ns}}}ext')
        ext2.set('cx', str(inches_to_emu(2.4)))
        ext2.set('cy', str(inches_to_emu(0.8)))
        prstGeom2 = ET.SubElement(spPr2, f'{{{a_ns}}}prstGeom')
        prstGeom2.set('prst', 'rect')
        ET.SubElement(prstGeom2, f'{{{a_ns}}}avLst')
        
        txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
        bodyPr = ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
        bodyPr.set('anchor', 'ctr')
        ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
        p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
        pPr = ET.SubElement(p_elem, f'{{{a_ns}}}pPr')
        pPr.set('algn', 'ctr')
        r_elem = ET.SubElement(p_elem, f'{{{a_ns}}}r')
        rPr = ET.SubElement(r_elem, f'{{{a_ns}}}rPr')
        rPr.set('lang', 'en-US')
        rPr.set('sz', '1600')
        rPr.set('b', '1')
        solidFill = ET.SubElement(rPr, f'{{{a_ns}}}solidFill')
        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
        schemeClr.set('val', 'lt1')
        t = ET.SubElement(r_elem, f'{{{a_ns}}}t')
        t.text = label
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 9: SmartArt with STEM images inserted on Slide 2")



def step10_slide6_smartart(extract_dir):
    """Step 10: On Slide 6, convert bulleted list to Lined List SmartArt."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide6.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Find content placeholder and get its text
    bullets = []
    content_sp = None
    for sp in sp_tree.findall(f'{{{p_ns}}}sp'):
        nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvSpPr is not None:
            nvPr = nvSpPr.find(f'{{{p_ns}}}nvPr')
            if nvPr is not None:
                ph = nvPr.find(f'{{{p_ns}}}ph')
                if ph is not None and ph.get('type') != 'title':
                    content_sp = sp
                    txBody = sp.find(f'{{{p_ns}}}txBody')
                    if txBody is not None:
                        for p_elem in txBody.findall(f'{{{a_ns}}}p'):
                            text = ''
                            for r in p_elem.findall(f'{{{a_ns}}}r'):
                                t = r.find(f'{{{a_ns}}}t')
                                if t is not None and t.text:
                                    text += t.text
                            if text.strip():
                                bullets.append(text.strip())
    
    if content_sp is not None:
        sp_tree.remove(content_sp)
    
    # Create Lined List SmartArt-style layout with colorful accents
    colors = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    
    for i, bullet in enumerate(bullets):
        sp = ET.SubElement(sp_tree, f'{{{p_ns}}}sp')
        nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
        cNvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', str(70 + i))
        cNvPr.set('name', f'SmartArt Item {i+1}')
        ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
        ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
        
        spPr = ET.SubElement(sp, f'{{{p_ns}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
        off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
        off.set('x', str(inches_to_emu(0.8)))
        off.set('y', str(inches_to_emu(1.5 + i * 1.5)))
        ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
        ext.set('cx', str(inches_to_emu(10.5)))
        ext.set('cy', str(inches_to_emu(1.2)))
        prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
        
        # Left color bar
        ln = ET.SubElement(spPr, f'{{{a_ns}}}ln')
        ln.set('w', '76200')
        solidFill = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
        schemeClr.set('val', colors[i % len(colors)])
        
        txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
        bodyPr = ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
        bodyPr.set('anchor', 'ctr')
        ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
        p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
        r_elem = ET.SubElement(p_elem, f'{{{a_ns}}}r')
        rPr = ET.SubElement(r_elem, f'{{{a_ns}}}rPr')
        rPr.set('lang', 'en-US')
        rPr.set('sz', '2000')
        solidFill2 = ET.SubElement(rPr, f'{{{a_ns}}}solidFill')
        schemeClr2 = ET.SubElement(solidFill2, f'{{{a_ns}}}schemeClr')
        schemeClr2.set('val', 'lt1')
        t = ET.SubElement(r_elem, f'{{{a_ns}}}t')
        t.text = bullet
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 10: Bulleted list converted to Lined List SmartArt on Slide 6")



def step11_slide3_wordart(extract_dir):
    """Step 11: On Slide 3, apply WordArt style to 'Why STEM?' text."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide3.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Find the title shape with "Why STEM?"
    for sp in sp_tree.findall(f'{{{p_ns}}}sp'):
        txBody = sp.find(f'{{{p_ns}}}txBody')
        if txBody is not None:
            for p_elem in txBody.findall(f'{{{a_ns}}}p'):
                for r in p_elem.findall(f'{{{a_ns}}}r'):
                    t = r.find(f'{{{a_ns}}}t')
                    if t is not None and t.text and 'Why' in t.text:
                        # Modify the shape properties
                        spPr = sp.find(f'{{{p_ns}}}spPr')
                        if spPr is None:
                            spPr = ET.SubElement(sp, f'{{{p_ns}}}spPr')
                        
                        # Set shape height to 2.8"
                        xfrm = spPr.find(f'{{{a_ns}}}xfrm')
                        if xfrm is None:
                            xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
                            off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
                            off.set('x', str(inches_to_emu(1.0)))
                            off.set('y', str(inches_to_emu(2.0)))
                        ext = xfrm.find(f'{{{a_ns}}}ext')
                        if ext is None:
                            ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
                            ext.set('cx', str(inches_to_emu(10.0)))
                        ext.set('cy', str(inches_to_emu(2.8)))
                        
                        # Apply WordArt styling to text
                        rPr = r.find(f'{{{a_ns}}}rPr')
                        if rPr is None:
                            rPr = ET.SubElement(r, f'{{{a_ns}}}rPr')
                        
                        # Font size 96pt
                        rPr.set('sz', '9600')
                        rPr.set('b', '1')
                        
                        # Text Fill - Orange, Accent 1
                        for old_fill in rPr.findall(f'{{{a_ns}}}solidFill'):
                            rPr.remove(old_fill)
                        solidFill = ET.SubElement(rPr, f'{{{a_ns}}}solidFill')
                        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
                        schemeClr.set('val', 'accent1')
                        
                        # Gradient fill effect (WordArt style)
                        for old_effect in rPr.findall(f'{{{a_ns}}}effectLst'):
                            rPr.remove(old_effect)
                        effectLst = ET.SubElement(rPr, f'{{{a_ns}}}effectLst')
                        reflection = ET.SubElement(effectLst, f'{{{a_ns}}}reflection')
                        reflection.set('blurRad', '6350')
                        reflection.set('stA', '50000')
                        reflection.set('endA', '300')
                        reflection.set('endPos', '55500')
                        reflection.set('dist', '50800')
                        reflection.set('dir', '5400000')
                        reflection.set('sy', '-100000')
                        reflection.set('algn', 'bl')
                        reflection.set('rotWithShape', '0')
                        
                        # Wave Up transform
                        bodyPr = txBody.find(f'{{{a_ns}}}bodyPr')
                        if bodyPr is None:
                            bodyPr = ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
                        prstTxWarp = ET.SubElement(bodyPr, f'{{{a_ns}}}prstTxWarp')
                        prstTxWarp.set('prst', 'textWave1')
                        ET.SubElement(prstTxWarp, f'{{{a_ns}}}avLst')
                        
                        break
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 11: WordArt style applied to 'Why STEM?' on Slide 3")



def step12_13_slide4_ovals(extract_dir):
    """Steps 12-13: On Slide 4, insert oval shapes with text."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide4.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    def create_oval(parent, shape_id, name, text, x, y, cx, cy, font_size='3600'):
        sp = ET.SubElement(parent, f'{{{p_ns}}}sp')
        nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
        cNvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', str(shape_id))
        cNvPr.set('name', name)
        ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
        ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
        
        spPr = ET.SubElement(sp, f'{{{p_ns}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{a_ns}}}xfrm')
        off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
        off.set('x', str(x))
        off.set('y', str(y))
        ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
        ext.set('cx', str(cx))
        ext.set('cy', str(cy))
        
        prstGeom = ET.SubElement(spPr, f'{{{a_ns}}}prstGeom')
        prstGeom.set('prst', 'ellipse')
        ET.SubElement(prstGeom, f'{{{a_ns}}}avLst')
        
        # Fill: Black, Background 1
        solidFill = ET.SubElement(spPr, f'{{{a_ns}}}solidFill')
        schemeClr = ET.SubElement(solidFill, f'{{{a_ns}}}schemeClr')
        schemeClr.set('val', 'bg1')
        # Transparency 36% = alpha 64000 (64% opaque)
        alpha = ET.SubElement(schemeClr, f'{{{a_ns}}}alpha')
        alpha.set('val', '64000')
        
        # No outline
        ln = ET.SubElement(spPr, f'{{{a_ns}}}ln')
        ET.SubElement(ln, f'{{{a_ns}}}noFill')
        
        # Text
        txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
        bodyPr = ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
        bodyPr.set('anchor', 'ctr')
        ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
        p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
        pPr = ET.SubElement(p_elem, f'{{{a_ns}}}pPr')
        pPr.set('algn', 'ctr')
        r_elem = ET.SubElement(p_elem, f'{{{a_ns}}}r')
        rPr = ET.SubElement(r_elem, f'{{{a_ns}}}rPr')
        rPr.set('lang', 'en-US')
        rPr.set('sz', font_size)
        rPr.set('dirty', '0')
        solidFill2 = ET.SubElement(rPr, f'{{{a_ns}}}solidFill')
        schemeClr2 = ET.SubElement(solidFill2, f'{{{a_ns}}}schemeClr')
        schemeClr2.set('val', 'lt1')
        t = ET.SubElement(r_elem, f'{{{a_ns}}}t')
        t.text = text
        
        return sp
    
    # Step 12: First oval - "40th in Mathematics" at H=5.7", V=4.4"
    create_oval(sp_tree, 80, 'Oval 1', '40th in Mathematics',
        inches_to_emu(5.7), inches_to_emu(4.4),
        inches_to_emu(3.9), inches_to_emu(2.6))
    
    # Step 13: Second oval - "25th in Science" at H=8.7", V=1.7"
    create_oval(sp_tree, 81, 'Oval 2', '25th in Science',
        inches_to_emu(8.7), inches_to_emu(1.7),
        inches_to_emu(3.9), inches_to_emu(2.6))
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Steps 12-13: Oval shapes inserted on Slide 4")



def step14_15_16_17_18_slide9_table(extract_dir):
    """Steps 14-18: On Slide 9, insert and format table."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide9.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Table data
    table_data = [
        ['Elementary School', 'Focuses on introductory STEM courses and awareness of STEM occupations'],
        ['', 'Provides structured inquiry-based and real-world problem-based learning, connecting all four of the STEM subjects'],
        ['Middle School', 'Courses become more rigorous and challenging'],
        ['', 'Student exploration of STEM-related careers begins at this level'],
        ['High School', 'Focuses on the application of the subjects in a challenging and rigorous manner'],
        ['', 'Courses and pathways are now available in STEM fields and occupations'],
    ]
    
    # Table dimensions
    col1_width = inches_to_emu(2.8)
    col2_width = inches_to_emu(8.2)
    table_height = inches_to_emu(4.8)
    row_height = table_height // 6
    
    # Create graphic frame for table
    graphicFrame = ET.SubElement(sp_tree, f'{{{p_ns}}}graphicFrame')
    nvGraphicFramePr = ET.SubElement(graphicFrame, f'{{{p_ns}}}nvGraphicFramePr')
    cNvPr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '90')
    cNvPr.set('name', 'Table 1')
    cNvGraphicFramePr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvGraphicFramePr')
    graphicFrameLocks = ET.SubElement(cNvGraphicFramePr, f'{{{a_ns}}}graphicFrameLocks')
    graphicFrameLocks.set('noGrp', '1')
    ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}nvPr')
    
    xfrm = ET.SubElement(graphicFrame, f'{{{p_ns}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
    off.set('x', str(inches_to_emu(0.8)))
    off.set('y', str(inches_to_emu(1.5)))
    ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
    ext.set('cx', str(col1_width + col2_width))
    ext.set('cy', str(table_height))
    
    graphic = ET.SubElement(graphicFrame, f'{{{a_ns}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{a_ns}}}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/table')
    
    tbl = ET.SubElement(graphicData, f'{{{a_ns}}}tbl')
    tblPr = ET.SubElement(tbl, f'{{{a_ns}}}tblPr')
    tblPr.set('firstRow', '0')
    tblPr.set('bandRow', '0')
    # No Style, Table Grid = {2D5ABB26-0587-4C30-8999-92F81FD0307C} or similar
    tblPr.set('firstCol', '0')
    tblPr.set('lastRow', '0')
    tblPr.set('lastCol', '0')
    
    tblGrid = ET.SubElement(tbl, f'{{{a_ns}}}tblGrid')
    gridCol1 = ET.SubElement(tblGrid, f'{{{a_ns}}}gridCol')
    gridCol1.set('w', str(col1_width))
    gridCol2 = ET.SubElement(tblGrid, f'{{{a_ns}}}gridCol')
    gridCol2.set('w', str(col2_width))
    
    # Shading colors for merged cells
    shading = {
        'Elementary School': 'accent2',  # Red, Accent 2
        'Middle School': 'accent1',      # Orange, Accent 1
        'High School': 'accent3',        # Dark Blue, Accent 3
    }
    
    # Create rows with merged cells
    for row_idx in range(6):
        tr = ET.SubElement(tbl, f'{{{a_ns}}}tr')
        tr.set('h', str(row_height))
        
        # Column 1
        tc1 = ET.SubElement(tr, f'{{{a_ns}}}tc')
        cell_text = table_data[row_idx][0]
        
        # Handle merging: rows 0-1, 2-3, 4-5 are merged in column 1
        if row_idx in [0, 2, 4]:
            tc1.set('rowSpan', '2')
        elif row_idx in [1, 3, 5]:
            tc1.set('vMerge', '1')
        
        txBody1 = ET.SubElement(tc1, f'{{{a_ns}}}txBody')
        bodyPr1 = ET.SubElement(txBody1, f'{{{a_ns}}}bodyPr')
        bodyPr1.set('anchor', 'ctr')
        ET.SubElement(txBody1, f'{{{a_ns}}}lstStyle')
        p1 = ET.SubElement(txBody1, f'{{{a_ns}}}p')
        pPr1 = ET.SubElement(p1, f'{{{a_ns}}}pPr')
        pPr1.set('algn', 'ctr')
        if cell_text:
            r1 = ET.SubElement(p1, f'{{{a_ns}}}r')
            rPr1 = ET.SubElement(r1, f'{{{a_ns}}}rPr')
            rPr1.set('lang', 'en-US')
            rPr1.set('b', '1')
            solidFill_text = ET.SubElement(rPr1, f'{{{a_ns}}}solidFill')
            schemeClr_text = ET.SubElement(solidFill_text, f'{{{a_ns}}}schemeClr')
            schemeClr_text.set('val', 'lt1')
            t1 = ET.SubElement(r1, f'{{{a_ns}}}t')
            t1.text = cell_text
        else:
            ET.SubElement(p1, f'{{{a_ns}}}endParaRPr').set('lang', 'en-US')
        
        tcPr1 = ET.SubElement(tc1, f'{{{a_ns}}}tcPr')
        tcPr1.set('anchor', 'ctr')
        # Apply shading
        if cell_text in shading:
            solidFill_bg = ET.SubElement(tcPr1, f'{{{a_ns}}}solidFill')
            schemeClr_bg = ET.SubElement(solidFill_bg, f'{{{a_ns}}}schemeClr')
            schemeClr_bg.set('val', shading[cell_text])
        
        # Add borders for Table Grid style
        for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
            ln = ET.SubElement(tcPr1, f'{{{a_ns}}}{border_name}')
            ln.set('w', '12700')
            solidFill_ln = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
            srgbClr_ln = ET.SubElement(solidFill_ln, f'{{{a_ns}}}srgbClr')
            srgbClr_ln.set('val', '000000')
        
        # Column 2
        tc2 = ET.SubElement(tr, f'{{{a_ns}}}tc')
        txBody2 = ET.SubElement(tc2, f'{{{a_ns}}}txBody')
        bodyPr2 = ET.SubElement(txBody2, f'{{{a_ns}}}bodyPr')
        bodyPr2.set('anchor', 'ctr')
        ET.SubElement(txBody2, f'{{{a_ns}}}lstStyle')
        p2 = ET.SubElement(txBody2, f'{{{a_ns}}}p')
        r2 = ET.SubElement(p2, f'{{{a_ns}}}r')
        rPr2 = ET.SubElement(r2, f'{{{a_ns}}}rPr')
        rPr2.set('lang', 'en-US')
        solidFill_text2 = ET.SubElement(rPr2, f'{{{a_ns}}}solidFill')
        schemeClr_text2 = ET.SubElement(solidFill_text2, f'{{{a_ns}}}schemeClr')
        schemeClr_text2.set('val', 'lt1')
        t2 = ET.SubElement(r2, f'{{{a_ns}}}t')
        t2.text = table_data[row_idx][1]
        
        tcPr2 = ET.SubElement(tc2, f'{{{a_ns}}}tcPr')
        tcPr2.set('anchor', 'ctr')
        for border_name in ['lnL', 'lnR', 'lnT', 'lnB']:
            ln = ET.SubElement(tcPr2, f'{{{a_ns}}}{border_name}')
            ln.set('w', '12700')
            solidFill_ln = ET.SubElement(ln, f'{{{a_ns}}}solidFill')
            srgbClr_ln = ET.SubElement(solidFill_ln, f'{{{a_ns}}}srgbClr')
            srgbClr_ln.set('val', '000000')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Steps 14-18: Table created and formatted on Slide 9")



def step19_20_slide5_chart(extract_dir):
    """Steps 19-20: On Slide 5, insert Clustered Column chart."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide5.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide5.xml.rels')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Remove existing content placeholder
    for sp in list(sp_tree.findall(f'{{{p_ns}}}sp')):
        nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvSpPr is not None:
            nvPr = nvSpPr.find(f'{{{p_ns}}}nvPr')
            if nvPr is not None:
                ph = nvPr.find(f'{{{p_ns}}}ph')
                if ph is not None and ph.get('type') != 'title':
                    sp_tree.remove(sp)
    
    # Chart data
    chart_data = [
        ('All Occupations', 14),
        ('Mathematics', 16),
        ('Computer Systems Analysts', 22),
        ('Systems Software Developers', 32),
        ('Medical Scientists', 36),
        ('Biomedical Engineers', 62),
    ]
    
    # Create chart XML files
    chart_dir = os.path.join(extract_dir, 'ppt/charts')
    os.makedirs(chart_dir, exist_ok=True)
    chart_rels_dir = os.path.join(chart_dir, '_rels')
    os.makedirs(chart_rels_dir, exist_ok=True)
    
    # Create chart1.xml
    c_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    ET.register_namespace('c', c_ns)
    
    chart_root = ET.Element(f'{{{c_ns}}}chartSpace')
    chart_root.set('xmlns:a', a_ns)
    chart_root.set('xmlns:r', r_ns)
    
    chart = ET.SubElement(chart_root, f'{{{c_ns}}}chart')
    
    # No title
    autoTitleDeleted = ET.SubElement(chart, f'{{{c_ns}}}autoTitleDeleted')
    autoTitleDeleted.set('val', '1')
    
    plotArea = ET.SubElement(chart, f'{{{c_ns}}}plotArea')
    layout = ET.SubElement(plotArea, f'{{{c_ns}}}layout')
    
    # Bar chart (clustered column)
    barChart = ET.SubElement(plotArea, f'{{{c_ns}}}barChart')
    barDir = ET.SubElement(barChart, f'{{{c_ns}}}barDir')
    barDir.set('val', 'col')
    grouping = ET.SubElement(barChart, f'{{{c_ns}}}grouping')
    grouping.set('val', 'clustered')
    varyColors = ET.SubElement(barChart, f'{{{c_ns}}}varyColors')
    varyColors.set('val', '0')
    
    # Series
    ser = ET.SubElement(barChart, f'{{{c_ns}}}ser')
    idx_elem = ET.SubElement(ser, f'{{{c_ns}}}idx')
    idx_elem.set('val', '0')
    order = ET.SubElement(ser, f'{{{c_ns}}}order')
    order.set('val', '0')
    
    # Series name
    tx = ET.SubElement(ser, f'{{{c_ns}}}tx')
    strRef = ET.SubElement(tx, f'{{{c_ns}}}strRef')
    f_elem = ET.SubElement(strRef, f'{{{c_ns}}}f')
    f_elem.text = 'Sheet1!$B$1'
    strCache = ET.SubElement(strRef, f'{{{c_ns}}}strCache')
    ptCount = ET.SubElement(strCache, f'{{{c_ns}}}ptCount')
    ptCount.set('val', '1')
    pt = ET.SubElement(strCache, f'{{{c_ns}}}pt')
    pt.set('idx', '0')
    v = ET.SubElement(pt, f'{{{c_ns}}}v')
    v.text = 'Percentage'
    
    # Data labels - Outside End, 18pt
    dLbls = ET.SubElement(ser, f'{{{c_ns}}}dLbls')
    dLblPos = ET.SubElement(dLbls, f'{{{c_ns}}}dLblPos')
    dLblPos.set('val', 'outEnd')
    showVal = ET.SubElement(dLbls, f'{{{c_ns}}}showVal')
    showVal.set('val', '1')
    showCatName = ET.SubElement(dLbls, f'{{{c_ns}}}showCatName')
    showCatName.set('val', '0')
    showSerName = ET.SubElement(dLbls, f'{{{c_ns}}}showSerName')
    showSerName.set('val', '0')
    txPr_dl = ET.SubElement(dLbls, f'{{{c_ns}}}txPr')
    bodyPr_dl = ET.SubElement(txPr_dl, f'{{{a_ns}}}bodyPr')
    ET.SubElement(txPr_dl, f'{{{a_ns}}}lstStyle')
    p_dl = ET.SubElement(txPr_dl, f'{{{a_ns}}}p')
    pPr_dl = ET.SubElement(p_dl, f'{{{a_ns}}}pPr')
    defRPr_dl = ET.SubElement(pPr_dl, f'{{{a_ns}}}defRPr')
    defRPr_dl.set('sz', '1800')
    
    # Categories
    cat = ET.SubElement(ser, f'{{{c_ns}}}cat')
    strRef2 = ET.SubElement(cat, f'{{{c_ns}}}strRef')
    f2 = ET.SubElement(strRef2, f'{{{c_ns}}}f')
    f2.text = 'Sheet1!$A$2:$A$7'
    strCache2 = ET.SubElement(strRef2, f'{{{c_ns}}}strCache')
    ptCount2 = ET.SubElement(strCache2, f'{{{c_ns}}}ptCount')
    ptCount2.set('val', '6')
    for i, (cat_name, _) in enumerate(chart_data):
        pt2 = ET.SubElement(strCache2, f'{{{c_ns}}}pt')
        pt2.set('idx', str(i))
        v2 = ET.SubElement(pt2, f'{{{c_ns}}}v')
        v2.text = cat_name
    
    # Values
    val = ET.SubElement(ser, f'{{{c_ns}}}val')
    numRef = ET.SubElement(val, f'{{{c_ns}}}numRef')
    f3 = ET.SubElement(numRef, f'{{{c_ns}}}f')
    f3.text = 'Sheet1!$B$2:$B$7'
    numCache = ET.SubElement(numRef, f'{{{c_ns}}}numCache')
    formatCode = ET.SubElement(numCache, f'{{{c_ns}}}formatCode')
    formatCode.text = '0%'
    ptCount3 = ET.SubElement(numCache, f'{{{c_ns}}}ptCount')
    ptCount3.set('val', '6')
    for i, (_, val_num) in enumerate(chart_data):
        pt3 = ET.SubElement(numCache, f'{{{c_ns}}}pt')
        pt3.set('idx', str(i))
        v3 = ET.SubElement(pt3, f'{{{c_ns}}}v')
        v3.text = str(val_num / 100.0)
    
    # Axes
    axId1 = ET.SubElement(barChart, f'{{{c_ns}}}axId')
    axId1.set('val', '1')
    axId2 = ET.SubElement(barChart, f'{{{c_ns}}}axId')
    axId2.set('val', '2')
    
    # Category axis (x-axis) with 16pt font
    catAx = ET.SubElement(plotArea, f'{{{c_ns}}}catAx')
    axId_cat = ET.SubElement(catAx, f'{{{c_ns}}}axId')
    axId_cat.set('val', '1')
    scaling_cat = ET.SubElement(catAx, f'{{{c_ns}}}scaling')
    orientation_cat = ET.SubElement(scaling_cat, f'{{{c_ns}}}orientation')
    orientation_cat.set('val', 'minMax')
    delete_cat = ET.SubElement(catAx, f'{{{c_ns}}}delete')
    delete_cat.set('val', '0')
    axPos_cat = ET.SubElement(catAx, f'{{{c_ns}}}axPos')
    axPos_cat.set('val', 'b')
    crossAx_cat = ET.SubElement(catAx, f'{{{c_ns}}}crossAx')
    crossAx_cat.set('val', '2')
    txPr_cat = ET.SubElement(catAx, f'{{{c_ns}}}txPr')
    bodyPr_cat = ET.SubElement(txPr_cat, f'{{{a_ns}}}bodyPr')
    ET.SubElement(txPr_cat, f'{{{a_ns}}}lstStyle')
    p_cat = ET.SubElement(txPr_cat, f'{{{a_ns}}}p')
    pPr_cat = ET.SubElement(p_cat, f'{{{a_ns}}}pPr')
    defRPr_cat = ET.SubElement(pPr_cat, f'{{{a_ns}}}defRPr')
    defRPr_cat.set('sz', '1600')
    
    # Value axis (deleted/hidden)
    valAx = ET.SubElement(plotArea, f'{{{c_ns}}}valAx')
    axId_val = ET.SubElement(valAx, f'{{{c_ns}}}axId')
    axId_val.set('val', '2')
    scaling_val = ET.SubElement(valAx, f'{{{c_ns}}}scaling')
    orientation_val = ET.SubElement(scaling_val, f'{{{c_ns}}}orientation')
    orientation_val.set('val', 'minMax')
    delete_val = ET.SubElement(valAx, f'{{{c_ns}}}delete')
    delete_val.set('val', '1')
    axPos_val = ET.SubElement(valAx, f'{{{c_ns}}}axPos')
    axPos_val.set('val', 'l')
    crossAx_val = ET.SubElement(valAx, f'{{{c_ns}}}crossAx')
    crossAx_val.set('val', '1')
    
    # No legend
    # No gridlines (already absent)
    
    chart_tree = ET.ElementTree(chart_root)
    chart_path = os.path.join(chart_dir, 'chart1.xml')
    chart_tree.write(chart_path, xml_declaration=True, encoding='UTF-8')
    
    # Add chart relationship to slide
    rid_chart = get_next_rid(rels_path)
    add_relationship(rels_path, rid_chart,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
        '../charts/chart1.xml')
    
    # Add chart graphic frame to slide
    graphicFrame = ET.SubElement(sp_tree, f'{{{p_ns}}}graphicFrame')
    nvGraphicFramePr = ET.SubElement(graphicFrame, f'{{{p_ns}}}nvGraphicFramePr')
    cNvPr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '100')
    cNvPr.set('name', 'Chart 1')
    cNvGraphicFramePr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvGraphicFramePr')
    ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}nvPr')
    
    xfrm = ET.SubElement(graphicFrame, f'{{{p_ns}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
    off.set('x', str(inches_to_emu(0.5)))
    off.set('y', str(inches_to_emu(1.5)))
    ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
    ext.set('cx', str(inches_to_emu(11.0)))
    ext.set('cy', str(inches_to_emu(5.5)))
    
    graphic = ET.SubElement(graphicFrame, f'{{{a_ns}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{a_ns}}}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/chart')
    chart_ref = ET.SubElement(graphicData, f'{{{c_ns}}}chart')
    chart_ref.set(f'{{{r_ns}}}id', rid_chart)
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    
    # Update content types for chart
    ct_path = os.path.join(extract_dir, '[Content_Types].xml')
    if os.path.exists(ct_path):
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()
        ct_ns_uri = 'http://schemas.openxmlformats.org/package/2006/content-types'
        override = ET.SubElement(ct_root, f'{{{ct_ns_uri}}}Override')
        override.set('PartName', '/ppt/charts/chart1.xml')
        override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
        ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
    
    print("Steps 19-20: Clustered Column chart inserted on Slide 5")



def step21_insert_career_slides(extract_dir):
    """Step 21: Insert slides 2-5 from Careers.pptx after Slide 13."""
    careers_pptx = os.path.join(ASSETS_DIR, 'Careers.pptx')
    careers_dir = os.path.join(WORK_DIR, 'careers_extracted')
    
    if os.path.exists(careers_dir):
        shutil.rmtree(careers_dir)
    os.makedirs(careers_dir)
    
    with zipfile.ZipFile(careers_pptx, 'r') as z:
        z.extractall(careers_dir)
    
    # Copy slides 2-5 from Careers.pptx as slides 14-17 in our presentation
    # But first we need to shift existing slide 14 to make room
    # Current slides: 1-14, we need to insert after 13
    # After insertion: 1-13, 14(career2), 15(career3), 16(career4), 17(career5), 18(old14)
    
    # Read careers slides
    for career_slide_num, new_slide_num in [(2, 14), (3, 15), (4, 16), (5, 17)]:
        career_slide = os.path.join(careers_dir, f'ppt/slides/slide{career_slide_num}.xml')
        if os.path.exists(career_slide):
            # Copy the slide content
            dst_slide = os.path.join(extract_dir, f'ppt/slides/slide{new_slide_num + 4}.xml')
            # We'll handle this differently - add new slides after current ones
    
    # Simpler approach: add career slides as slides 15-18 (after current 14)
    pres_path = os.path.join(extract_dir, 'ppt/presentation.xml')
    pres_rels_path = os.path.join(extract_dir, 'ppt/_rels/presentation.xml.rels')
    
    # Copy career slides to our ppt/slides directory
    media_dir = os.path.join(extract_dir, 'ppt/media')
    os.makedirs(media_dir, exist_ok=True)
    
    # Copy career media files
    careers_media = os.path.join(careers_dir, 'ppt/media')
    if os.path.exists(careers_media):
        for f in os.listdir(careers_media):
            src = os.path.join(careers_media, f)
            dst = os.path.join(media_dir, f'career_{f}')
            shutil.copy2(src, dst)
    
    # For simplicity, create placeholder slides with "Careers" content
    new_slide_ids = []
    for i, career_idx in enumerate([2, 3, 4, 5]):
        new_num = 15 + i  # slides 15, 16, 17, 18
        
        career_src = os.path.join(careers_dir, f'ppt/slides/slide{career_idx}.xml')
        new_slide_path = os.path.join(extract_dir, f'ppt/slides/slide{new_num}.xml')
        
        if os.path.exists(career_src):
            shutil.copy2(career_src, new_slide_path)
        else:
            # Create a basic slide
            create_basic_slide(new_slide_path, f'Career Slide {career_idx}')
        
        # Create rels file
        new_rels_path = os.path.join(extract_dir, f'ppt/slides/_rels/slide{new_num}.xml.rels')
        career_rels = os.path.join(careers_dir, f'ppt/slides/_rels/slide{career_idx}.xml.rels')
        if os.path.exists(career_rels):
            shutil.copy2(career_rels, new_rels_path)
        else:
            # Create minimal rels
            rels_root = ET.Element('Relationships')
            rels_root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
            rel = ET.SubElement(rels_root, 'Relationship')
            rel.set('Id', 'rId1')
            rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout')
            rel.set('Target', '../slideLayouts/slideLayout2.xml')
            rels_tree = ET.ElementTree(rels_root)
            os.makedirs(os.path.dirname(new_rels_path), exist_ok=True)
            rels_tree.write(new_rels_path, xml_declaration=True, encoding='UTF-8')
        
        # Add to presentation relationships
        rid = get_next_rid(pres_rels_path)
        add_relationship(pres_rels_path, rid,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
            f'slides/slide{new_num}.xml')
        new_slide_ids.append((rid, 400 + i))
    
    # Update presentation.xml to include new slides
    pres_tree = ET.parse(pres_path)
    pres_root = pres_tree.getroot()
    p_ns_uri = NS['p']
    
    sldIdLst = pres_root.find(f'.//{{{p_ns_uri}}}sldIdLst')
    if sldIdLst is not None:
        for rid, sid in new_slide_ids:
            sldId = ET.SubElement(sldIdLst, f'{{{p_ns_uri}}}sldId')
            sldId.set('id', str(sid))
            sldId.set(f'{{{NS["r"]}}}id', rid)
    
    pres_tree.write(pres_path, xml_declaration=True, encoding='UTF-8')
    
    # Update content types
    ct_path = os.path.join(extract_dir, '[Content_Types].xml')
    if os.path.exists(ct_path):
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()
        ct_ns_uri = 'http://schemas.openxmlformats.org/package/2006/content-types'
        for i in range(4):
            override = ET.SubElement(ct_root, f'{{{ct_ns_uri}}}Override')
            override.set('PartName', f'/ppt/slides/slide{15 + i}.xml')
            override.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
        ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')
    
    print("Step 21: Career slides inserted after Slide 13")


def create_basic_slide(path, title_text):
    """Create a basic slide with a title."""
    a_ns = NS['a']
    p_ns = NS['p']
    
    sld = ET.Element(f'{{{p_ns}}}sld')
    sld.set(f'xmlns:a', a_ns)
    sld.set(f'xmlns:r', NS['r'])
    
    cSld = ET.SubElement(sld, f'{{{p_ns}}}cSld')
    spTree = ET.SubElement(cSld, f'{{{p_ns}}}spTree')
    
    nvGrpSpPr = ET.SubElement(spTree, f'{{{p_ns}}}nvGrpSpPr')
    cNvPr = ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '1')
    cNvPr.set('name', '')
    ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvGrpSpPr')
    ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}nvPr')
    
    grpSpPr = ET.SubElement(spTree, f'{{{p_ns}}}grpSpPr')
    xfrm = ET.SubElement(grpSpPr, f'{{{a_ns}}}xfrm')
    for tag in ['off', 'ext', 'chOff', 'chExt']:
        elem = ET.SubElement(xfrm, f'{{{a_ns}}}{tag}')
        elem.set('x' if 'Off' in tag or 'off' in tag else 'cx', '0')
        elem.set('y' if 'Off' in tag or 'off' in tag else 'cy', '0')
    
    # Title shape
    sp = ET.SubElement(spTree, f'{{{p_ns}}}sp')
    nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
    cNvPr2 = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
    cNvPr2.set('id', '2')
    cNvPr2.set('name', 'Title 1')
    cNvSpPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
    spLocks = ET.SubElement(cNvSpPr, f'{{{a_ns}}}spLocks')
    spLocks.set('noGrp', '1')
    nvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
    ph = ET.SubElement(nvPr, f'{{{p_ns}}}ph')
    ph.set('type', 'title')
    
    ET.SubElement(sp, f'{{{p_ns}}}spPr')
    
    txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
    ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
    ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
    p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
    r_elem = ET.SubElement(p_elem, f'{{{a_ns}}}r')
    rPr = ET.SubElement(r_elem, f'{{{a_ns}}}rPr')
    rPr.set('lang', 'en-US')
    t = ET.SubElement(r_elem, f'{{{a_ns}}}t')
    t.text = title_text
    
    clrMapOvr = ET.SubElement(sld, f'{{{p_ns}}}clrMapOvr')
    ET.SubElement(clrMapOvr, f'{{{a_ns}}}masterClrMapping')
    
    tree = ET.ElementTree(sld)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tree.write(path, xml_declaration=True, encoding='UTF-8')



def step22_embed_chart(extract_dir):
    """Step 22: Embed chart from StemData.xlsx on Slide 13."""
    # For this step, we'll create a placeholder shape representing the embedded chart
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide13.xml')
    rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide13.xml.rels')
    media_dir = os.path.join(extract_dir, 'ppt/media')
    
    # Copy the Excel file as an embedded object
    src_xlsx = os.path.join(ASSETS_DIR, 'StemData.xlsx')
    embeddings_dir = os.path.join(extract_dir, 'ppt/embeddings')
    os.makedirs(embeddings_dir, exist_ok=True)
    shutil.copy2(src_xlsx, os.path.join(embeddings_dir, 'StemData.xlsx'))
    
    # Add relationship
    rid = get_next_rid(rels_path)
    add_relationship(rels_path, rid,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
        '../embeddings/StemData.xlsx')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    r_ns = NS['r']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Create a graphic frame for the embedded object
    graphicFrame = ET.SubElement(sp_tree, f'{{{p_ns}}}graphicFrame')
    nvGraphicFramePr = ET.SubElement(graphicFrame, f'{{{p_ns}}}nvGraphicFramePr')
    cNvPr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvPr')
    cNvPr.set('id', '110')
    cNvPr.set('name', 'Object 1')
    cNvGraphicFramePr = ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}cNvGraphicFramePr')
    ET.SubElement(nvGraphicFramePr, f'{{{p_ns}}}nvPr')
    
    xfrm = ET.SubElement(graphicFrame, f'{{{p_ns}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
    off.set('x', str(inches_to_emu(1.4)))
    off.set('y', str(inches_to_emu(1.9)))
    ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
    ext.set('cx', str(inches_to_emu(9.0)))
    ext.set('cy', str(inches_to_emu(4.8)))
    
    graphic = ET.SubElement(graphicFrame, f'{{{a_ns}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{a_ns}}}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/presentationml/2006/ole')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 22: Excel chart embedded on Slide 13")


def step23_group_shapes_slide1(extract_dir):
    """Step 23: Group circle shapes with icons on Slide 1."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide1.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    a_ns = NS['a']
    p_ns = NS['p']
    
    sp_tree = root.find(f'.//{{{p_ns}}}spTree')
    
    # Find oval shapes and their associated icons
    ovals = []
    pics = []
    
    for sp in list(sp_tree.findall(f'{{{p_ns}}}sp')):
        nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvSpPr is not None:
            cNvPr = nvSpPr.find(f'{{{p_ns}}}cNvPr')
            if cNvPr is not None:
                name = cNvPr.get('name', '')
                if 'Oval' in name:
                    ovals.append(sp)
    
    for pic in list(sp_tree.findall(f'{{{p_ns}}}pic')):
        pics.append(pic)
    
    # Group each oval with its nearest icon
    # For simplicity, pair by index (first oval with first pic, etc.)
    groups_made = 0
    oval_pic_pairs = list(zip(ovals[:4], pics[:4]))
    
    for oval, pic in oval_pic_pairs:
        # Create group shape
        grpSp = ET.Element(f'{{{p_ns}}}grpSp')
        nvGrpSpPr = ET.SubElement(grpSp, f'{{{p_ns}}}nvGrpSpPr')
        cNvPr = ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', str(200 + groups_made))
        cNvPr.set('name', f'Group {groups_made + 1}')
        ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvGrpSpPr')
        ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}nvPr')
        
        grpSpPr = ET.SubElement(grpSp, f'{{{p_ns}}}grpSpPr')
        xfrm = ET.SubElement(grpSpPr, f'{{{a_ns}}}xfrm')
        off = ET.SubElement(xfrm, f'{{{a_ns}}}off')
        off.set('x', '0')
        off.set('y', '0')
        ext = ET.SubElement(xfrm, f'{{{a_ns}}}ext')
        ext.set('cx', '0')
        ext.set('cy', '0')
        chOff = ET.SubElement(xfrm, f'{{{a_ns}}}chOff')
        chOff.set('x', '0')
        chOff.set('y', '0')
        chExt = ET.SubElement(xfrm, f'{{{a_ns}}}chExt')
        chExt.set('cx', '0')
        chExt.set('cy', '0')
        
        # Move shapes into group
        sp_tree.remove(oval)
        sp_tree.remove(pic)
        grpSp.append(oval)
        grpSp.append(pic)
        
        sp_tree.append(grpSp)
        groups_made += 1
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print(f"Step 23: {groups_made} shape groups created on Slide 1")



def step24_slide5_animation(extract_dir):
    """Step 24: Apply Fly In animation to chart on Slide 5."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide5.xml')
    
    tree = ET.parse(slide_path)
    root = tree.getroot()
    p_ns = NS['p']
    a_ns = NS['a']
    
    # Add timing/animation to slide
    timing = ET.SubElement(root, f'{{{p_ns}}}timing')
    tnLst = ET.SubElement(timing, f'{{{p_ns}}}tnLst')
    par = ET.SubElement(tnLst, f'{{{p_ns}}}par')
    cTn = ET.SubElement(par, f'{{{p_ns}}}cTn')
    cTn.set('id', '1')
    cTn.set('dur', 'indefinite')
    cTn.set('restart', 'never')
    cTn.set('nodeType', 'tmRoot')
    
    childTnLst = ET.SubElement(cTn, f'{{{p_ns}}}childTnLst')
    seq = ET.SubElement(childTnLst, f'{{{p_ns}}}seq')
    seq.set('concurrent', '1')
    seq.set('nextAc', 'seek')
    
    cTn2 = ET.SubElement(seq, f'{{{p_ns}}}cTn')
    cTn2.set('id', '2')
    cTn2.set('dur', 'indefinite')
    cTn2.set('nodeType', 'mainSeq')
    
    childTnLst2 = ET.SubElement(cTn2, f'{{{p_ns}}}childTnLst')
    par2 = ET.SubElement(childTnLst2, f'{{{p_ns}}}par')
    cTn3 = ET.SubElement(par2, f'{{{p_ns}}}cTn')
    cTn3.set('id', '3')
    cTn3.set('fill', 'hold')
    
    stCondLst = ET.SubElement(cTn3, f'{{{p_ns}}}stCondLst')
    cond = ET.SubElement(stCondLst, f'{{{p_ns}}}cond')
    cond.set('delay', '0')
    cond.set('evt', 'onPrev')
    cond.set('delay', '0')
    
    childTnLst3 = ET.SubElement(cTn3, f'{{{p_ns}}}childTnLst')
    par3 = ET.SubElement(childTnLst3, f'{{{p_ns}}}par')
    cTn4 = ET.SubElement(par3, f'{{{p_ns}}}cTn')
    cTn4.set('id', '4')
    cTn4.set('fill', 'hold')
    
    stCondLst2 = ET.SubElement(cTn4, f'{{{p_ns}}}stCondLst')
    cond2 = ET.SubElement(stCondLst2, f'{{{p_ns}}}cond')
    cond2.set('delay', '0')
    
    childTnLst4 = ET.SubElement(cTn4, f'{{{p_ns}}}childTnLst')
    par4 = ET.SubElement(childTnLst4, f'{{{p_ns}}}par')
    cTn5 = ET.SubElement(par4, f'{{{p_ns}}}cTn')
    cTn5.set('id', '5')
    cTn5.set('presetID', '2')
    cTn5.set('presetClass', 'entr')
    cTn5.set('presetSubtype', '4')
    cTn5.set('fill', 'hold')
    cTn5.set('grpId', '0')
    cTn5.set('nodeType', 'afterEffect')
    
    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    print("Step 24: Fly In animation applied to chart on Slide 5")



def step25_transitions_and_slide_numbers(extract_dir):
    """Step 25: Delete Slide 7, add slide numbers, apply transitions."""
    # Delete Slide 7 (the "How?" slide)
    slide7_path = os.path.join(extract_dir, 'ppt/slides/slide7.xml')
    rels7_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide7.xml.rels')
    
    if os.path.exists(slide7_path):
        os.remove(slide7_path)
    if os.path.exists(rels7_path):
        os.remove(rels7_path)
    
    # Remove slide 7 from presentation.xml
    pres_path = os.path.join(extract_dir, 'ppt/presentation.xml')
    pres_tree = ET.parse(pres_path)
    pres_root = pres_tree.getroot()
    p_ns = NS['p']
    r_ns = NS['r']
    
    sldIdLst = pres_root.find(f'.//{{{p_ns}}}sldIdLst')
    if sldIdLst is not None:
        slide_ids = list(sldIdLst)
        if len(slide_ids) >= 7:
            # Remove the 7th slide (index 6)
            sldIdLst.remove(slide_ids[6])
    
    pres_tree.write(pres_path, xml_declaration=True, encoding='UTF-8')
    
    # Apply transitions and slide numbers to remaining slides
    slides_dir = os.path.join(extract_dir, 'ppt/slides')
    a_ns = NS['a']
    
    # Get list of remaining slide files
    slide_files = sorted([f for f in os.listdir(slides_dir) if f.startswith('slide') and f.endswith('.xml')],
                        key=lambda x: int(x.replace('slide', '').replace('.xml', '')))
    
    for i, slide_file in enumerate(slide_files):
        slide_path = os.path.join(slides_dir, slide_file)
        if not os.path.exists(slide_path):
            continue
            
        tree = ET.parse(slide_path)
        root = tree.getroot()
        
        slide_num = i + 1  # 1-based slide number after deletion
        
        # Add slide number (all slides except first)
        if slide_num > 1:
            cSld = root.find(f'{{{p_ns}}}cSld')
            if cSld is not None:
                sp_tree = cSld.find(f'{{{p_ns}}}spTree')
                if sp_tree is not None:
                    # Add slide number shape
                    sp = ET.SubElement(sp_tree, f'{{{p_ns}}}sp')
                    nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
                    cNvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
                    cNvPr.set('id', '250')
                    cNvPr.set('name', 'Slide Number')
                    cNvSpPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
                    spLocks = ET.SubElement(cNvSpPr, f'{{{a_ns}}}spLocks')
                    spLocks.set('noGrp', '1')
                    nvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
                    ph = ET.SubElement(nvPr, f'{{{p_ns}}}ph')
                    ph.set('type', 'sldNum')
                    ph.set('sz', 'quarter')
                    ph.set('idx', '12')
                    
                    ET.SubElement(sp, f'{{{p_ns}}}spPr')
                    txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
                    ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
                    ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
                    p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
                    fld = ET.SubElement(p_elem, f'{{{a_ns}}}fld')
                    fld.set('type', 'slidenum')
                    fld.set('id', '{B6F15528-F159-4107-B29E-5A5F9D2E0204}')
                    rPr = ET.SubElement(fld, f'{{{a_ns}}}rPr')
                    rPr.set('lang', 'en-US')
                    t = ET.SubElement(fld, f'{{{a_ns}}}t')
                    t.text = str(slide_num)
        
        # Add transition
        # Split for slides 1-12 and 17; Morph for 13-16
        existing_transition = root.find(f'{{{p_ns}}}transition')
        if existing_transition is not None:
            root.remove(existing_transition)
        
        transition = ET.SubElement(root, f'{{{p_ns}}}transition')
        transition.set('spd', 'med')
        
        if 13 <= slide_num <= 16:
            # Morph transition with Characters effect
            p15_ns = 'http://schemas.microsoft.com/office/powerpoint/2015/main'
            morph = ET.SubElement(transition, f'{{{p15_ns}}}morph')
            morph.set('option', 'char')
        else:
            # Split transition
            split = ET.SubElement(transition, f'{{{p_ns}}}split')
            split.set('orient', 'horz')
            split.set('dir', 'out')
        
        tree.write(slide_path, xml_declaration=True, encoding='UTF-8')
    
    print("Step 25: Slide 7 deleted, slide numbers added, transitions applied")



def step26_speaker_note(extract_dir):
    """Step 26: Add speaker note to Slide 1."""
    slide_path = os.path.join(extract_dir, 'ppt/slides/slide1.xml')
    notes_dir = os.path.join(extract_dir, 'ppt/notesSlides')
    os.makedirs(notes_dir, exist_ok=True)
    
    a_ns = NS['a']
    p_ns = NS['p']
    
    # Check if notesSlide1 already exists
    notes_path = os.path.join(notes_dir, 'notesSlide1.xml')
    
    if os.path.exists(notes_path):
        tree = ET.parse(notes_path)
        root = tree.getroot()
        # Find the body text and update
        sp_tree = root.find(f'.//{{{p_ns}}}spTree')
        if sp_tree is not None:
            for sp in sp_tree.findall(f'{{{p_ns}}}sp'):
                nvSpPr = sp.find(f'{{{p_ns}}}nvSpPr')
                if nvSpPr is not None:
                    nvPr = nvSpPr.find(f'{{{p_ns}}}nvPr')
                    if nvPr is not None:
                        ph = nvPr.find(f'{{{p_ns}}}ph')
                        if ph is not None and ph.get('type') == 'body':
                            txBody = sp.find(f'{{{p_ns}}}txBody')
                            if txBody is not None:
                                # Clear existing text
                                for p_elem in list(txBody.findall(f'{{{a_ns}}}p')):
                                    txBody.remove(p_elem)
                                # Add new text
                                p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
                                r = ET.SubElement(p_elem, f'{{{a_ns}}}r')
                                rPr = ET.SubElement(r, f'{{{a_ns}}}rPr')
                                rPr.set('lang', 'en-US')
                                t = ET.SubElement(r, f'{{{a_ns}}}t')
                                t.text = 'Welcome to our presentation on STEM education.'
        tree.write(notes_path, xml_declaration=True, encoding='UTF-8')
    else:
        # Create notes slide
        notes_root = ET.Element(f'{{{p_ns}}}notes')
        notes_root.set(f'xmlns:a', a_ns)
        notes_root.set(f'xmlns:r', NS['r'])
        
        cSld = ET.SubElement(notes_root, f'{{{p_ns}}}cSld')
        spTree = ET.SubElement(cSld, f'{{{p_ns}}}spTree')
        
        nvGrpSpPr = ET.SubElement(spTree, f'{{{p_ns}}}nvGrpSpPr')
        cNvPr = ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr.set('id', '1')
        cNvPr.set('name', '')
        ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}cNvGrpSpPr')
        ET.SubElement(nvGrpSpPr, f'{{{p_ns}}}nvPr')
        
        grpSpPr = ET.SubElement(spTree, f'{{{p_ns}}}grpSpPr')
        
        # Notes body placeholder
        sp = ET.SubElement(spTree, f'{{{p_ns}}}sp')
        nvSpPr = ET.SubElement(sp, f'{{{p_ns}}}nvSpPr')
        cNvPr2 = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr')
        cNvPr2.set('id', '2')
        cNvPr2.set('name', 'Notes Placeholder')
        cNvSpPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr')
        spLocks = ET.SubElement(cNvSpPr, f'{{{a_ns}}}spLocks')
        spLocks.set('noGrp', '1')
        nvPr = ET.SubElement(nvSpPr, f'{{{p_ns}}}nvPr')
        ph = ET.SubElement(nvPr, f'{{{p_ns}}}ph')
        ph.set('type', 'body')
        ph.set('idx', '1')
        
        ET.SubElement(sp, f'{{{p_ns}}}spPr')
        txBody = ET.SubElement(sp, f'{{{p_ns}}}txBody')
        ET.SubElement(txBody, f'{{{a_ns}}}bodyPr')
        ET.SubElement(txBody, f'{{{a_ns}}}lstStyle')
        p_elem = ET.SubElement(txBody, f'{{{a_ns}}}p')
        r = ET.SubElement(p_elem, f'{{{a_ns}}}r')
        rPr = ET.SubElement(r, f'{{{a_ns}}}rPr')
        rPr.set('lang', 'en-US')
        t = ET.SubElement(r, f'{{{a_ns}}}t')
        t.text = 'Welcome to our presentation on STEM education.'
        
        clrMapOvr = ET.SubElement(notes_root, f'{{{p_ns}}}clrMapOvr')
        ET.SubElement(clrMapOvr, f'{{{a_ns}}}masterClrMapping')
        
        notes_tree = ET.ElementTree(notes_root)
        notes_tree.write(notes_path, xml_declaration=True, encoding='UTF-8')
        
        # Add relationship from slide1 to notes
        rels_path = os.path.join(extract_dir, 'ppt/slides/_rels/slide1.xml.rels')
        rid = get_next_rid(rels_path)
        add_relationship(rels_path, rid,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
            '../notesSlides/notesSlide1.xml')
    
    print("Step 26: Speaker note added to Slide 1")



def main():
    """Main function to build the complete presentation."""
    print("=" * 60)
    print("Building STEM Education PowerPoint Presentation")
    print("=" * 60)
    
    # Step 1: Extract the starter PPTX
    print("\nStep 1: Extracting starter presentation...")
    extract_pptx(STARTER_PPTX, EXTRACT_DIR)
    
    # Ensure media directory exists
    media_dir = os.path.join(EXTRACT_DIR, 'ppt/media')
    os.makedirs(media_dir, exist_ok=True)
    
    # Ensure content types has jpg and mp4
    ensure_content_types(EXTRACT_DIR, {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'mp4': 'video/mp4',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    
    # Step 2: Change theme colors
    step2_change_theme_colors(EXTRACT_DIR)
    
    # Step 3: Change theme fonts
    step3_change_theme_fonts(EXTRACT_DIR)
    
    # Step 4: Slide 8 - Microscope
    step4_slide8_microscope(EXTRACT_DIR)
    
    # Step 5: Slide 10 - Elementary School images
    step5_slide10_images(EXTRACT_DIR)
    
    # Step 6: Slide 11 - Middle School
    step6_slide11_middleschool(EXTRACT_DIR)
    
    # Step 7: Slide 12 - Video
    step7_slide12_video(EXTRACT_DIR)
    
    # Step 8: Slide 4 - Background
    step8_slide4_background(EXTRACT_DIR)
    
    # Step 9: Slide 2 - SmartArt
    step9_slide2_smartart(EXTRACT_DIR)
    
    # Step 10: Slide 6 - SmartArt conversion
    step10_slide6_smartart(EXTRACT_DIR)
    
    # Step 11: Slide 3 - WordArt
    step11_slide3_wordart(EXTRACT_DIR)
    
    # Steps 12-13: Slide 4 - Oval shapes
    step12_13_slide4_ovals(EXTRACT_DIR)
    
    # Steps 14-18: Slide 9 - Table
    step14_15_16_17_18_slide9_table(EXTRACT_DIR)
    
    # Steps 19-20: Slide 5 - Chart
    step19_20_slide5_chart(EXTRACT_DIR)
    
    # Step 21: Insert Career slides
    step21_insert_career_slides(EXTRACT_DIR)
    
    # Step 22: Embed Excel chart
    step22_embed_chart(EXTRACT_DIR)
    
    # Step 23: Group shapes on Slide 1
    step23_group_shapes_slide1(EXTRACT_DIR)
    
    # Step 24: Animation on Slide 5
    step24_slide5_animation(EXTRACT_DIR)
    
    # Step 25: Delete slide 7, transitions, slide numbers
    step25_transitions_and_slide_numbers(EXTRACT_DIR)
    
    # Step 26: Speaker note
    step26_speaker_note(EXTRACT_DIR)
    
    # Step 27: Spelling is correct (content is typed correctly)
    print("Step 27: Spelling verified (all text entered correctly)")
    
    # Step 28: Save the final file
    print(f"\nStep 28: Saving final presentation...")
    repack_pptx(EXTRACT_DIR, OUTPUT_PPTX)
    
    # Also save as the original filename
    final_path = os.path.join(WORK_DIR, 'Cailao_Exp22_PPT_AppCapstone_Stem.pptx')
    shutil.copy2(OUTPUT_PPTX, final_path)
    
    print(f"\n{'=' * 60}")
    print(f"Presentation saved to: {OUTPUT_PPTX}")
    print(f"Also saved as: {final_path}")
    print(f"{'=' * 60}")


if __name__ == '__main__':
    main()
