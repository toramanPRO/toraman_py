from hashlib import sha256
import datetime
import random
import os
import string
import zipfile

import Levenshtein
from lxml import etree
import regex

__version__ = '0.1.1'

nsmap = {'toraman': 'https://cat.toraman.pro'}


class BilingualFile:
    def __init__(self, file_path):
        self.hyperlinks = []
        self.images = []
        self.paragraphs = []
        self.tags = []

        self.xml_root = etree.parse(file_path).getroot()
        self.t_nsmap = self.xml_root.nsmap

        self.file_type = self.xml_root.find('toraman:source_file', self.t_nsmap).attrib['type']
        self.file_name = self.xml_root.find('toraman:source_file', self.t_nsmap).attrib['name']
        self.nsmap = self.xml_root.find('toraman:source_file', self.t_nsmap)[0][0].nsmap

        for xml_paragraph in self.xml_root.find('toraman:paragraphs', self.t_nsmap):
            paragraph_no = xml_paragraph.attrib['no']
            current_paragraph = []
            for xml_segment in xml_paragraph.findall('toraman:segment', self.t_nsmap):
                current_paragraph.append([xml_segment[0],
                                          xml_segment[1],
                                          xml_segment[2],
                                          int(paragraph_no),
                                          int(xml_segment.attrib['no'])])

            self.paragraphs.append(current_paragraph)

        if self.xml_root.find('toraman:tags', self.t_nsmap) is not None:
            for xml_tag in self.xml_root.find('toraman:tags', self.t_nsmap):
                self.tags.append(xml_tag.text)

        if self.xml_root.find('toraman:hyperlinks', self.t_nsmap) is not None:
            for xml_hl in self.xml_root.find('toraman:hyperlinks', self.t_nsmap):
                self.hyperlinks.append(xml_hl.text)

        if self.xml_root.find('toraman:images', self.t_nsmap) is not None:
            for xml_image in self.xml_root.find('toraman:images', self.t_nsmap):
                self.images.append(xml_image.text)

    def generate_target_translation(self, source_file_path, output_directory):
        sha256_hash = sha256()
        buffer_size = 5 * 1048576  # 5 MB
        sf = open(source_file_path, 'rb')
        while True:
            data = sf.read(buffer_size)
            if data:
                sha256_hash.update(data)
            else:
                break
        sf.close()
        sha256_hash = sha256_hash.hexdigest()

        sf_sha256 = self.xml_root.find('toraman:source_file', self.t_nsmap).attrib['sha256']

        assert sha256_hash == sf_sha256, 'SHA256 hash of the file does not match that of the source file'

        target_paragraphs = []
        for xml_paragraph in self.xml_root[0]:
            target_paragraph = []
            for xml_segment in xml_paragraph:
                if xml_segment.tag == '{{{0}}}segment'.format(self.t_nsmap['toraman']):
                    if len(xml_segment[2]) == 0:
                        for source_elem in xml_segment[0]:
                            target_paragraph.append(source_elem)
                    else:
                        for target_elem in xml_segment[2]:
                            target_paragraph.append(target_elem)
                else:
                    for source_elem in xml_segment:
                        target_paragraph.append(source_elem)

            target_paragraphs.append(target_paragraph)

        if self.file_type == 'docx':

            final_paragraphs = []

            for target_paragraph in target_paragraphs:
                active_ftags = []
                final_paragraph = [etree.Element('{{{0}}}r'.format(self.nsmap['w']))]
                start_new_run = False
                for sub_elem in target_paragraph:
                    if start_new_run:
                        if len(final_paragraph[-1]) > 0:
                            final_paragraph.append(etree.Element('{{{0}}}r'.format(self.nsmap['w'])))
                        start_new_run = False
                        if active_ftags:
                            final_paragraph[-1].append(etree.Element('{{{0}}}rPr'.format(self.nsmap['w'])))
                            for active_ftag_no in reversed(active_ftags):
                                if int(active_ftag_no) <= len(self.tags):
                                    active_ftag = etree.fromstring(self.tags[int(active_ftag_no)-1])
                                    for prop in active_ftag:
                                        if final_paragraph[-1][-1].find(prop.tag) is None:
                                            final_paragraph[-1][-1].append(prop)

                    if sub_elem.tag == '{{{0}}}text'.format(self.t_nsmap['toraman']):
                        if (final_paragraph[-1].find('w:t', self.nsmap) is not None
                                and final_paragraph[-1][-1].tag == '{{{0}}}t'.format(self.nsmap['w'])):
                            final_paragraph[-1][-1].text += sub_elem.text
                        else:
                            final_paragraph[-1].append(etree.Element('{{{0}}}t'.format(self.nsmap['w'])))
                            final_paragraph[-1][-1].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            final_paragraph[-1][-1].text = sub_elem.text
                    elif sub_elem.tag == '{{{0}}}tag'.format(self.t_nsmap['toraman']):
                        start_new_run = True
                        if sub_elem.attrib['type'] == 'beginning':
                            if sub_elem.attrib['no'] not in active_ftags:
                                active_ftags.append(sub_elem.attrib['no'])
                        else:
                            if sub_elem.attrib['no'] in active_ftags:
                                active_ftags.remove(sub_elem.attrib['no'])
                    elif sub_elem.tag == '{{{0}}}tab'.format(self.t_nsmap['toraman']):
                        final_paragraph[-1].append(etree.Element('{{{0}}}tab'.format(self.nsmap['w'])))
                    elif sub_elem.tag == '{{{0}}}br'.format(self.t_nsmap['toraman']):
                        if 'type' in sub_elem.attrib and sub_elem.attrib['type'] == 'page':
                            if len(final_paragraph[-1]) > 0:
                                final_paragraph.append(etree.Element('{{{0}}}r'.format(self.nsmap['w'])))
                            final_paragraph[-1].append(etree.Element('{{{0}}}br'.format(self.nsmap['w'])))
                            final_paragraph[-1][-1].attrib['type'] = 'page'
                        else:
                            final_paragraph[-1].append(etree.Element('{{{0}}}br'.format(self.nsmap['w'])))
                    elif sub_elem.tag == '{{{0}}}image'.format(self.t_nsmap['toraman']):
                        if int(sub_elem.attrib['no']) <= len(self.images):
                            if len(final_paragraph[-1]) > 0:
                                final_paragraph.append(etree.Element('{{{0}}}r'.format(self.nsmap['w'])))
                            final_paragraph[-1].append(etree.fromstring(self.images[int(sub_elem.attrib['no'])-1]))
                            final_paragraph.append(etree.Element('{{{0}}}r'.format(self.nsmap['w'])))

                if len(final_paragraph[-1]) == 0:
                    final_paragraph = final_paragraph[:-1]

                final_paragraphs.append(final_paragraph)

            for internal_file in self.xml_root[-1]:
                internal_file = internal_file[0]
                for paragraph in internal_file.findall('.//w:p', self.nsmap):
                    for paragraph_placeholder in paragraph.findall('toraman:paragraph', self.t_nsmap):
                        current_elem_i = paragraph.index(paragraph_placeholder)
                        paragraph.remove(paragraph_placeholder)
                        for final_run in final_paragraphs[int(paragraph_placeholder.attrib['no'])-1]:
                            if len(final_run) == 0:
                                pass
                            elif len(final_run) == 1 and final_run[0].tag.endswith('rPr'):
                                pass
                            else:
                                paragraph.insert(current_elem_i,
                                                etree.fromstring(etree.tostring(final_run)))
                                current_elem_i += 1

            with zipfile.ZipFile(source_file_path) as zf:
                for name in zf.namelist():
                    zf.extract(name, os.path.join(output_directory, '.temp'))

            for internal_file in self.xml_root[-1]:
                etree.ElementTree(internal_file[0]).write(os.path.join(output_directory,
                                                                       '.temp',
                                                                       internal_file.attrib['internal_path']),
                                                          encoding='UTF-8',
                                                          xml_declaration=True)

            to_zip = []
            for root, dir, files in os.walk(os.path.join(output_directory, '.temp')):
                for name in files:
                    to_zip.append(os.path.join(root, name))

            with zipfile.ZipFile(os.path.join(output_directory, self.file_name), 'w') as target_zf:
                for name in to_zip:
                    target_zf.write(name, name[len(os.path.join(output_directory, '.temp')):])

    def save(self, output_directory):
        self.xml_root.getroottree().write(os.path.join(output_directory, self.file_name) + '.xml',
                                          encoding='UTF-8',
                                          xml_declaration=True)

    def update_segment(self, segment_status, segment_target, paragraph_no, segment_no):

        assert type(segment_target) == etree._Element or type(segment_target) == str

        xml_segment = self.xml_root[0][paragraph_no - 1].find('toraman:segment[@no="{0}"]'.format(segment_no),
                                                              self.t_nsmap)
        sub_p_id = xml_segment.getparent().findall('toraman:segment', self.t_nsmap).index(xml_segment)
        segment = self.paragraphs[paragraph_no - 1][sub_p_id]

        xml_segment[1].text = segment_status
        xml_segment[2] = etree.Element('{{{0}}}target'.format(self.t_nsmap['toraman']))

        if type(segment_target) == str:
            xml_segment[2].append(etree.Element('{{{0}}}text'.format(self.t_nsmap['toraman'])))
            xml_segment[2][0].text = segment_target
        else:
            for sub_elem in segment_target:
                xml_segment[2].append(sub_elem)

        segment[1] = xml_segment[1]
        segment[2] = xml_segment[2]

        self.paragraphs[paragraph_no - 1][sub_p_id] = segment


class SourceFile:
    def __init__(self, file_path, list_of_abbreviations=None):
        self.file_type = ''
        self.file_name = ''
        self.hyperlinks = []
        self.images = []
        self.master_files = []
        self.nsmap = None
        self.paragraphs = []
        self.sha256_hash = sha256()
        self.tags = []
        self.t_nsmap = nsmap

        buffer_size = 5 * 1048576  # 5 MB
        sf = open(file_path, 'rb')
        while True:
            data = sf.read(buffer_size)
            if data:
                self.sha256_hash.update(data)
            else:
                break
        sf.close()

        self.sha256_hash = self.sha256_hash.hexdigest()

        self.file_name = file_path.replace('\\', '/').split('/')[-1]

        if file_path.lower().endswith('.docx'):

            def extract_p_run(r_element, p_element, paragraph_continues):
                toraman_run = etree.Element('{{{0}}}run'.format(self.t_nsmap['toraman']), nsmap=self.t_nsmap)
                run_properties = r_element.find('w:rPr', self.nsmap)
                if len(r_element) == 1:
                    if (r_element[0].tag == '{{{0}}}br'.format(self.nsmap['w']) and
                        b'type="page"' in etree.tostring(r_element[0])):
                        paragraph_continues = False

                        return paragraph_continues

                    elif (not paragraph_continues and run_properties is not None):

                        return paragraph_continues                    

                if run_properties is not None:
                    for run_property in run_properties:
                        if ('{{{0}}}val'.format(self.nsmap['w']) in run_property.attrib
                                and run_property.attrib['{{{0}}}val'.format(self.nsmap['w'])].lower() == 'false'):
                            run_properties.remove(run_property)
                        elif run_property.tag == '{{{0}}}lang'.format(self.nsmap['w']):
                            run_properties.remove(run_property)
                        elif run_property.tag == '{{{0}}}noProof'.format(self.nsmap['w']):
                            run_properties.remove(run_property)
                    if len(run_properties) > 0:
                        run_properties = [run_properties,
                                          etree.Element('{{{0}}}rPr'.format(self.nsmap['w']), nsmap=self.nsmap)]
                        for run_property in run_properties[0]:
                            run_properties[1].append(run_property)
                        run_properties = etree.tostring(run_properties[1])
                    else:
                        run_properties = None

                for sub_r_element in r_element:
                    if sub_r_element.tag == '{{{0}}}t'.format(self.nsmap['w']):
                        if (toraman_run.find('toraman:text', self.t_nsmap) is not None
                           and toraman_run[-1].tag == '{{{0}}}text'.format(self.t_nsmap['toraman'])):
                            toraman_run[-1].text += sub_r_element.text
                        else:
                            toraman_run.append(etree.Element('{{{0}}}text'.format(self.t_nsmap['toraman']),
                                                             nsmap=self.t_nsmap))
                            toraman_run[-1].text = sub_r_element.text
                    elif sub_r_element.tag == '{{{0}}}tab'.format(self.nsmap['w']):
                        toraman_run.append(etree.Element('{{{0}}}tab'.format(self.t_nsmap['toraman']),
                                                         nsmap=self.t_nsmap))
                    elif sub_r_element.tag == '{{{0}}}br'.format(self.nsmap['w']):
                        if ('{{{0}}}type'.format(self.nsmap['w']) in sub_r_element.attrib
                                and sub_r_element.attrib['{{{0}}}type'.format(self.nsmap['w'])] == 'page'):
                            toraman_run.append(etree.Element('{{{0}}}br'.format(self.t_nsmap['toraman']),
                                                             type='page', nsmap=self.t_nsmap))
                        else:
                            toraman_run.append(etree.Element('{{{0}}}br'.format(self.t_nsmap['toraman']),
                                                             nsmap=self.t_nsmap))
                    elif sub_r_element.tag == '{{{0}}}drawing'.format(self.nsmap['w']):
                        self.images.append(etree.tostring(sub_r_element))
                        toraman_run.append(etree.Element('{{{0}}}image'.format(self.t_nsmap['toraman']),
                                                         no=str(len(self.images)),
                                                         nsmap=self.t_nsmap))
                    elif sub_r_element.tag == '{{{0}}}rPr'.format(self.nsmap['w']):
                        pass

                if paragraph_continues:
                    self.paragraphs[-1].append([toraman_run, run_properties])
                    p_element.remove(r_element)
                else:
                    self.paragraphs.append([[toraman_run, run_properties]])
                    p_element.replace(r_element,
                                      etree.Element('{{{0}}}paragraph'.format(self.t_nsmap['toraman']),
                                                    no=str(len(self.paragraphs)),
                                                    nsmap=self.t_nsmap))
                    paragraph_continues = True

                return paragraph_continues

            sf = zipfile.ZipFile(file_path)
            for zip_child in sf.namelist():
                if 'word/document.xml' in zip_child:
                    self.master_files.append([zip_child, sf.open(zip_child)])
                elif 'word/document2.xml' in zip_child:
                    self.master_files.append([zip_child, sf.open(zip_child)])
            sf.close()

            assert self.master_files
            self.file_type = 'docx'

            for master_file in self.master_files:
                master_file[1] = etree.parse(master_file[1])

                master_file[1] = master_file[1].getroot()
                self.nsmap = master_file[1].nsmap

                placeholders = []
                for paragraph_element in master_file[1].findall('.//w:p', self.nsmap):
                    if paragraph_element.find('.//w:t', self.nsmap) is None:
                        placeholders.append(paragraph_element)
                        placeholder_xml = etree.Element('{{{0}}}paragraph_placeholder'.format(self.t_nsmap['toraman']),
                                                        no=str(len(placeholders)),
                                                        nsmap=self.t_nsmap)
                        paragraph_element.getparent().replace(paragraph_element,
                                                              placeholder_xml)

                for paragraph_element in master_file[1].xpath('w:body/w:p|w:body/w:tbl/w:tr', namespaces=self.nsmap):
                    add_to_last_paragraph = False
                    for sub_element in paragraph_element:
                        if (sub_element.tag == '{{{0}}}r'.format(self.nsmap['w'])
                                and sub_element.find('{{{0}}}AlternateContent'.format(self.nsmap['mc'])) is not None):
                            add_to_last_paragraph = False
                            tb_elements = sub_element.findall('.//{{{0}}}txbxContent'.format(self.nsmap['w']))
                            if len(tb_elements) == 1 or len(tb_elements) == 2:
                                for tb_paragraph_element in tb_elements[0].findall('w:p', self.nsmap):
                                    for tb_sub_element in tb_paragraph_element:
                                        if tb_sub_element.tag == '{{{0}}}r'.format(self.nsmap['w']):
                                            add_to_last_paragraph = extract_p_run(tb_sub_element,
                                                                                  tb_paragraph_element,
                                                                                  add_to_last_paragraph)
                                        elif tb_sub_element.tag == '{{{0}}}pPr'.format(self.nsmap['w']):
                                            pass
                                        else:
                                            tb_paragraph_element.remove(tb_sub_element)
                                if len(tb_elements) == 2:
                                    tb_elements.append(etree.fromstring(etree.tostring(tb_elements[0])))
                                    tb_elements[1].getparent().replace(tb_elements[1], tb_elements[2])
                            add_to_last_paragraph = False
                        elif sub_element.tag == '{{{0}}}r'.format(self.nsmap['w']):
                            add_to_last_paragraph = extract_p_run(sub_element,
                                                                  paragraph_element,
                                                                  add_to_last_paragraph)
                        elif sub_element.tag == '{{{0}}}hyperlink'.format(self.nsmap['w']):
                            add_to_last_paragraph = extract_p_run(sub_element.find('w:r', self.nsmap),
                                                                  sub_element,
                                                                  add_to_last_paragraph)
                            self.paragraphs[-1][-1].append(etree.tostring(sub_element))
                            paragraph_element.remove(sub_element)
                        elif sub_element.tag == '{{{0}}}pPr'.format(self.nsmap['w']):
                            pass
                        elif sub_element.tag == '{{{0}}}tc'.format(self.nsmap['w']):
                            for tc_p_element in sub_element.findall('w:p', self.nsmap):
                                for sub_tc_p_element in tc_p_element:
                                    if sub_tc_p_element.tag == '{{{0}}}r'.format(self.nsmap['w']):
                                        add_to_last_paragraph = extract_p_run(sub_tc_p_element,
                                                                              tc_p_element,
                                                                              add_to_last_paragraph)
                        else:
                            paragraph_element.remove(sub_element)

                for paragraph_placeholder in master_file[1].findall('.//toraman:paragraph_placeholder', self.t_nsmap):
                    paragraph_element = placeholders[int(paragraph_placeholder.attrib['no']) - 1]
                    paragraph_placeholder.getparent().replace(paragraph_placeholder, paragraph_element)

        for paragraph_index in range(len(self.paragraphs)):
            organised_paragraph = etree.Element('{{{0}}}OrganisedParagraph'.format(self.t_nsmap['toraman']),
                                                nsmap=self.t_nsmap)
            for run in self.paragraphs[paragraph_index]:
                if len(run) == 2 and run[1] is not None:
                    if run[1] not in self.tags:
                        self.tags.append(run[1])
                    toraman_tag_template = etree.tostring(etree.Element('{{{0}}}tag'.format(self.t_nsmap['toraman']),
                                                                        no=str(self.tags.index(run[1])+1),
                                                                        nsmap=self.t_nsmap))

                    toraman_tag = etree.fromstring(toraman_tag_template)
                    toraman_tag.attrib['type'] = 'beginning'
                    organised_paragraph.append(toraman_tag)

                    for run_element in run[0]:
                        organised_paragraph.append(run_element)

                    toraman_tag = etree.fromstring(toraman_tag_template)
                    toraman_tag.attrib['type'] = 'end'
                    organised_paragraph.append(toraman_tag)

                elif len(run) == 3:
                    if run[1] is not None:
                        if run[1] not in self.tags:
                            self.tags.append(run[1])
                        toraman_tag_template = etree.tostring(etree.Element('{{{0}}}tag'.format(self.t_nsmap['toraman']),
                                                                            no=str(self.tags.index(run[1])+1),
                                                                            nsmap=self.t_nsmap))

                        toraman_tag = etree.fromstring(toraman_tag_template)
                        toraman_tag.attrib['type'] = 'beginning'
                        organised_paragraph.append(toraman_tag)

                    if run[2] is not None:
                        if run[2] not in self.hyperlinks:
                            self.hyperlinks.append(run[2])
                        toraman_link_template = etree.tostring(etree.Element('{{{0}}}link'.format(self.t_nsmap['toraman']),
                                                                             no=str(self.hyperlinks.index(run[2])+1),
                                                                             nsmap=self.t_nsmap))

                        toraman_link = etree.fromstring(toraman_link_template)
                        toraman_link.attrib['type'] = 'beginning'
                        organised_paragraph.append(toraman_tag)

                    for run_element in run[0]:
                        organised_paragraph.append(run_element)

                    if run[2] is not None:
                        toraman_link = etree.fromstring(toraman_link_template)
                        toraman_link.attrib['type'] = 'end'
                        organised_paragraph.append(toraman_tag)

                    if run[1] is not None:
                        toraman_tag = etree.fromstring(toraman_tag_template)
                        toraman_tag.attrib['type'] = 'end'
                        organised_paragraph.append(toraman_tag)

                else:
                    for run_element in run[0]:
                        organised_paragraph.append(run_element)

            for toraman_tag_end in organised_paragraph.findall('toraman:tag[@type="end"]', self.t_nsmap):
                if (toraman_tag_end.getnext() is not None
                        and toraman_tag_end.tag == toraman_tag_end.getnext().tag
                        and toraman_tag_end.attrib['no'] == toraman_tag_end.getnext().attrib['no']
                        and toraman_tag_end.getnext().attrib['type'] == 'beginning'):
                    organised_paragraph.remove(toraman_tag_end.getnext())
                    organised_paragraph.remove(toraman_tag_end)

            placeholders = ['placeholder_to_keep_segment_going',
                            'placeholder_to_end_segment']
            while placeholders[0] in str(etree.tostring(organised_paragraph)):
                placeholders[0] += random.choice(string.ascii_letters)
            while placeholders[1] in str(etree.tostring(organised_paragraph)):
                placeholders[1] += random.choice(string.ascii_letters)

            _regex = regex.compile(r'(\s+|^)'
                                   r'(\p{Lu}\p{L}{0,3})'
                                   r'(\.+)'
                                   r'(\s+|$)')
            for toraman_t in organised_paragraph.findall('toraman:text', self.t_nsmap):
                mid_sentence_punctuation = []
                for _hit in regex.findall(_regex, toraman_t.text):
                    if _hit is not None:
                        _hit = list(_hit)
                        mid_sentence_punctuation.append(_hit)
                for to_be_replaced in mid_sentence_punctuation:
                    toraman_t.text = regex.sub(''.join(to_be_replaced),
                                               ''.join((to_be_replaced[0],
                                                        to_be_replaced[1],
                                                        to_be_replaced[2],
                                                        placeholders[0],
                                                        to_be_replaced[3])),
                                               toraman_t.text,
                                               1)

            if list_of_abbreviations:
                list_of_abbreviations = '|'.join(list_of_abbreviations)
                _regex = regex.compile(r'(\s+|^)({0})(\.+)(\s+|$)'.format(list_of_abbreviations))
                for toraman_t in organised_paragraph.findall('toraman:text', self.t_nsmap):
                    mid_sentence_punctuation = []
                    for _hit in regex.findall(_regex, toraman_t.text):
                        if _hit is not None:
                            mid_sentence_punctuation.append(_hit)
                    for to_be_replaced in mid_sentence_punctuation:
                        toraman_t.text = regex.sub(''.join(to_be_replaced),
                                                   ''.join((to_be_replaced[0],
                                                            to_be_replaced[1],
                                                            to_be_replaced[2],
                                                            placeholders[0],
                                                            to_be_replaced[3])),
                                                   toraman_t.text,
                                                   1)

            _regex = regex.compile(r'(\s+|^)'
                                   r'([\p{Lu}\p{L}]+)'
                                   r'([\.\!\?\:]+)'
                                   r'(\s+|$)')
            for toraman_t in organised_paragraph.findall('toraman:text', self.t_nsmap):
                end_sentence_punctuation = []
                for _hit in regex.findall(_regex, toraman_t.text):
                    if _hit is not None:
                        end_sentence_punctuation.append(_hit)
                for to_be_replaced in end_sentence_punctuation:
                    toraman_t.text = regex.sub(''.join(to_be_replaced) + '(?!placeholder)',
                                               ''.join((to_be_replaced[0],
                                                        to_be_replaced[1],
                                                        to_be_replaced[2],
                                                        placeholders[1],
                                                        to_be_replaced[3])),
                                               toraman_t.text,
                                               1)
                toraman_t.text = regex.sub(placeholders[0],
                                           '',
                                           toraman_t.text)

            organised_paragraph = [[], organised_paragraph]
            organised_paragraph[0].append(etree.Element('{{{0}}}source'.format(self.t_nsmap['toraman']),
                                                        nsmap=self.t_nsmap))
            for toraman_element in organised_paragraph[1]:
                if toraman_element.tag == '{{{0}}}text'.format(self.t_nsmap['toraman']):
                    _text = toraman_element.text.split(placeholders[1])
                    for _text_i in range(len(_text)):
                        if _text_i != 0:
                            organised_paragraph[0].append(etree.Element('{{{0}}}source'.format(self.t_nsmap['toraman']),
                                                                        nsmap=self.t_nsmap))
                        organised_paragraph[0][-1].append(etree.Element('{{{0}}}text'.format(self.t_nsmap['toraman']),
                                                            nsmap=self.t_nsmap))
                        organised_paragraph[0][-1][-1].text = _text[_text_i]
                elif (toraman_element.tag == '{{{0}}}br'.format(self.t_nsmap['toraman'])
                        and 'type' in toraman_element.attrib 
                        and toraman_element.attrib['type'] == 'page'):
                    if len(organised_paragraph[0][-1]) == 0:
                        organised_paragraph[0] = organised_paragraph[0][:-1]
                    organised_paragraph[0].append(etree.Element('{{{0}}}non-text-segment'.format(self.t_nsmap['toraman']),
                                                                nsmap=self.t_nsmap))
                    organised_paragraph[0][-1].append(toraman_element)
                elif toraman_element.tag == '{{{0}}}br'.format(self.t_nsmap['toraman']):
                    if (len(organised_paragraph[0][-1]) == 0
                            or (len(organised_paragraph[0][-1]) == 1
                            and organised_paragraph[0][-1][-1].tag == '{{{0}}}text'.format(self.t_nsmap['toraman'])
                            and organised_paragraph[0][-1][-1].text is '')):
                        organised_paragraph[0] = organised_paragraph[0][:-1]
                        organised_paragraph[0].append(etree.Element('{{{0}}}non-text-segment'.format(self.t_nsmap['toraman']),
                                                                    nsmap=self.t_nsmap))

                    organised_paragraph[0][-1].append(toraman_element)
                    organised_paragraph[0].append(etree.Element('{{{0}}}source'.format(self.t_nsmap['toraman']),
                                                        nsmap=self.t_nsmap))

                else:
                    organised_paragraph[0][-1].append(toraman_element)

            if (len(organised_paragraph[0][-1]) == 0
                    or (len(organised_paragraph[0][-1]) == 1
                    and organised_paragraph[0][-1][-1].tag == '{{{0}}}text'.format(self.t_nsmap['toraman'])
                    and organised_paragraph[0][-1][-1].text is '')):
                organised_paragraph[0] = organised_paragraph[0][:-1]

            self.paragraphs[paragraph_index] = organised_paragraph[0]

    def write_bilingual_file(self, output_directory):
        bilingual_file = etree.Element('{{{0}}}bilingual_file'.format(self.t_nsmap['toraman']), nsmap=self.t_nsmap)

        etree.SubElement(bilingual_file, '{{{0}}}paragraphs'.format(self.t_nsmap['toraman']))

        _counter = 0
        for p_i in range(len(self.paragraphs)):
            new_p_element = etree.Element('{{{0}}}paragraph'.format(self.t_nsmap['toraman']), no=str(p_i + 1))
            for segment_element in self.paragraphs[p_i]:
                if segment_element.tag == '{{{0}}}source'.format(self.t_nsmap['toraman']):
                    _counter += 1
                    new_s_element = etree.Element('{{{0}}}segment'.format(self.t_nsmap['toraman']), no=str(_counter))
                    new_s_element.append(segment_element)
                    new_s_element.append(etree.Element('{{{0}}}status'.format(self.t_nsmap['toraman'])))
                    new_s_element.append(etree.Element('{{{0}}}target'.format(self.t_nsmap['toraman'])))
                else:
                    new_s_element = segment_element
                new_p_element.append(new_s_element)
            bilingual_file[-1].append(new_p_element)

        if self.images:
            bilingual_file.append(etree.Element('{{{0}}}images'.format(self.t_nsmap['toraman'])))
            for i_i in range(len(self.images)):
                new_i_element = etree.Element('{{{0}}}image'.format(self.t_nsmap['toraman']), no=str(i_i + 1))
                new_i_element.text = self.images[i_i]
                bilingual_file[-1].append(new_i_element)

        if self.tags:
            bilingual_file.append(etree.Element('{{{0}}}tags'.format(self.t_nsmap['toraman'])))
            for t_i in range(len(self.tags)):
                new_t_element = etree.Element('{{{0}}}tag'.format(self.t_nsmap['toraman']), no=str(t_i + 1))
                new_t_element.text = self.tags[t_i]
                bilingual_file[-1].append(new_t_element)

        if self.hyperlinks:
            bilingual_file.append(etree.Element('{{{0}}}hyperlinks'.format(self.t_nsmap['toraman'])))
            for h_i in range(len(self.hyperlinks)):
                new_h_element = etree.Element('{{{0}}}hyperlink'.format(self.t_nsmap['toraman']), no=str(h_i + 1))
                new_h_element.text = self.hyperlinks[h_i]
                bilingual_file[-1].append(new_h_element)

        bilingual_file.append(etree.Element('{{{0}}}source_file'.format(self.t_nsmap['toraman']),
                                            sha256=self.sha256_hash,
                                            name=self.file_name,
                                            type=self.file_type))
        for master_file in self.master_files:
            new_sf_element = etree.Element('{{{0}}}internal_file'.format(self.t_nsmap['toraman']),
                                           internal_path=master_file[0])
            new_sf_element.append(master_file[1])
            bilingual_file[-1].append(new_sf_element)

        bilingual_file.getroottree().write(os.path.join(output_directory, self.file_name) + '.xml',
                                           encoding='UTF-8',
                                           xml_declaration=True)


class TranslationMemory():
    def __init__(self, tm_path, source_language, target_language):
        self.tm_path = tm_path
        if os.path.exists(tm_path):
            self.translation_memory = etree.parse(tm_path).getroot()

            self.source_language = self.translation_memory[0].attrib['srclang']
            self.target_language = self.translation_memory[0].attrib['trgtlang']

        else:
            self.source_language = source_language
            self.target_language = target_language

            self.translation_memory = etree.Element('{{{0}}}tm'.format(nsmap['toraman']), nsmap=nsmap)
            self.translation_memory.attrib['version'] = '0.0.1'
            self.translation_memory.append(etree.Element('{{{0}}}header'.format(nsmap['toraman'])))
            self.translation_memory[0].attrib['creationtool'] = 'toraman'
            self.translation_memory[0].attrib['creationtoolversion'] = __version__
            self.translation_memory[0].attrib['creationdate'] = datetime.datetime.utcnow().strftime(r'%Y%M%dT%H%M%SZ')
            self.translation_memory[0].attrib['datatype'] = 'PlainText'
            self.translation_memory[0].attrib['segtype'] = 'sentence'
            self.translation_memory[0].attrib['adminlang'] = 'en'
            self.translation_memory[0].attrib['srclang'] = self.source_language
            self.translation_memory[0].attrib['trgtlang'] = self.target_language
            self.translation_memory.append(etree.Element('{{{0}}}body'.format(nsmap['toraman'])))

            self.translation_memory.getroottree().write(self.tm_path,
                                           encoding='UTF-8',
                                           xml_declaration=True)

    def lookup(self, source_segment):
        _segment_hits = []

        source_segment = etree.tostring(source_segment)

        for translation_unit in self.translation_memory[1]:
            saved_source_segment = etree.tostring(translation_unit[0])
            if Levenshtein.ratio(saved_source_segment, source_segment) >= 0.70:
                saved_target_segment = etree.tostring(translation_unit[1])
                _segment_hits.append((Levenshtein.ratio(saved_source_segment, source_segment),
                                    etree.fromstring(saved_source_segment),
                                    etree.fromstring(saved_target_segment)))
        else:
            _segment_hits.sort(reverse=True)

        return _segment_hits

    def submit_segment(self, source_segment, target_segment):
        source_segment = etree.tostring(source_segment)
        target_segment = etree.tostring(target_segment)

        for translation_unit in self.translation_memory[1]:
            if etree.tostring(translation_unit[0]) == source_segment:
                translation_unit[1] = etree.fromstring(target_segment)
                translation_unit.attrib['changedate'] = datetime.datetime.utcnow().strftime(r'%Y%M%dT%H%M%SZ')
                break
        else:
            translation_unit = etree.Element('{{{0}}}tu'.format(nsmap['toraman']))
            translation_unit.attrib['creationdate'] = datetime.datetime.utcnow().strftime(r'%Y%M%dT%H%M%SZ')
            translation_unit.append(etree.fromstring(source_segment))
            translation_unit.append(etree.fromstring(target_segment))

            self.translation_memory[1].append(translation_unit)

        self.translation_memory.getroottree().write(self.tm_path,
                                           encoding='UTF-8',
                                           xml_declaration=True)
