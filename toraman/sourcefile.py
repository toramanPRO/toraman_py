from hashlib import sha256
import os, random, string, zipfile

from lxml import etree
import regex

from .utils import nsmap


class SourceFile:
    def __init__(self, file_path, list_of_abbreviations=None):
        self.file_type = ''
        self.file_name = ''
        self.hyperlinks = []
        self.images = []
        self.master_files = []
        self.miscellaneous_tags = []
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
                                add_to_last_paragraph = False
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

        elif file_path.lower().endswith('.odt'):
            sf = zipfile.ZipFile(file_path)
            for zip_child in sf.namelist():
                if ('content.xml' in zip_child
                or 'styles.xml' in zip_child):
                    self.master_files.append([zip_child, sf.open(zip_child)])
            sf.close()

            assert self.master_files
            self.file_type = 'odt'

            def extract_span(child_element, parent_element, paragraph_continues):
                run_properties = child_element.attrib['{{{0}}}style-name'.format(self.nsmap['text'])]
                if not paragraph_continues:
                    self.paragraphs.append([[etree.Element('{{{0}}}run'.format(nsmap['toraman'])), run_properties]])

                    parent_element.replace(child_element, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                            no=str(len(self.paragraphs))))

                    paragraph_continues = True
                else:
                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman'])), run_properties])

                    parent_element.remove(child_element)

                if child_element.text is not None:
                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                    self.paragraphs[-1][-1][0][-1].text = child_element.text

                for span_child in child_element:
                    if span_child.tag == '{{{0}}}frame'.format(self.nsmap['draw']):
                        image_copy = child_element.__deepcopy__(True)
                        image_copy.tail = None
                        self.images.append(etree.tostring(image_copy))
                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}image'.format(self.t_nsmap['toraman']),
                                                        no=str(len(self.images)),
                                                        nsmap=self.t_nsmap))
                    elif span_child.tag == '{{{0}}}line-break'.format(self.nsmap['text']):
                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}br'.format(nsmap['toraman'])))
                    elif span_child.tag == '{{{0}}}tab'.format(self.nsmap['text']):
                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}tab'.format(nsmap['toraman'])))
                    elif span_child.tag == '{{{0}}}s'.format(self.nsmap['text']):
                        if (len(self.paragraphs[-1][-1][0]) == 0
                        or self.paragraphs[-1][-1][0][-1].tag != '{{{0}}}text'.format(nsmap['toraman'])):
                            self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                            self.paragraphs[-1][-1][0][-1].text = ' '
                        else:
                            self.paragraphs[-1][-1][0][-1].text += ' '

                    if span_child.tail is not None:
                        if (len(self.paragraphs[-1][-1][0]) == 0
                        or self.paragraphs[-1][-1][0][-1].tag != '{{{0}}}text'.format(nsmap['toraman'])):
                            self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                            self.paragraphs[-1][-1][0][-1].text = span_child.tail
                        else:
                            self.paragraphs[-1][-1][0][-1].text += span_child.tail

            for master_file in self.master_files:
                master_file[1] = etree.parse(master_file[1])

                master_file[1] = master_file[1].getroot()
                self.nsmap = master_file[1].nsmap

                for paragraph_element in master_file[1].xpath('office:body/office:text/text:p|office:body/office:text/table:table//text:p|office:master-styles/style:master-page/style:header/text:p|office:master-styles/style:master-page/style:footer/text:p', namespaces=self.nsmap):
                    paragraph_continues = False

                    if paragraph_element.text is not None:
                        self.paragraphs.append([[etree.Element('{{{0}}}run'.format(nsmap['toraman']))]])

                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                        self.paragraphs[-1][-1][0][-1].text = paragraph_element.text
                        paragraph_element.text = None
                        paragraph_element.insert(0, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                    no=str(len(self.paragraphs))))
                        paragraph_continues = True

                    for paragraph_child in paragraph_element:

                        if paragraph_child.tag == '{{{0}}}frame'.format(self.nsmap['draw']):
                            if paragraph_child[0].tag == '{{{0}}}text-box'.format(self.nsmap['draw']):
                                for tb_paragraph_element in paragraph_child[0].findall('text:p', self.nsmap):
                                    paragraph_continues = False
                                    if tb_paragraph_element.text is not None:
                                        self.paragraphs.append([[etree.Element('{{{0}}}run'.format(nsmap['toraman']))]])

                                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                        self.paragraphs[-1][-1][0][-1].text = tb_paragraph_element.text
                                        tb_paragraph_element.text = None
                                        tb_paragraph_element.insert(0, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                                    no=str(len(self.paragraphs))))
                                        paragraph_continues = True

                                    for tb_paragraph_child in tb_paragraph_element:
                                        if tb_paragraph_child.tag == '{{{0}}}frame'.format(self.nsmap['draw']):
                                            if (tb_paragraph_child.tail is not None
                                            or paragraph_element.find('toraman:paragraph', nsmap) is not None
                                            or paragraph_element.find('text:a', self.nsmap) is not None
                                            or paragraph_element.find('text:span', self.nsmap) is not None
                                            or paragraph_element.find('text.s', self.nsmap) is not None):
                                                image_copy = tb_paragraph_child.__deepcopy__(True)
                                                image_copy.tail = None
                                                self.images.append(etree.tostring(image_copy))
                                                if paragraph_continues:
                                                    if len(self.paragraphs[-1][-1]) != 1:
                                                        self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                                else:
                                                    if len(self.paragraphs[-1][-1]) != 1:
                                                        self.paragraphs[-1].append([[etree.Element('{{{0}}}run'.format(nsmap['toraman']))]])
                                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}image'.format(self.t_nsmap['toraman']),
                                                                                no=str(len(self.images)),
                                                                                nsmap=self.t_nsmap))

                                        elif tb_paragraph_child.tag == '{{{0}}}span'.format(self.nsmap['text']):
                                            extract_span(tb_paragraph_child, tb_paragraph_element, paragraph_continues)

                                            paragraph_continues = True

                                        if tb_paragraph_child.tail is not None:
                                            if len(self.paragraphs[-1][-1]) == 1:
                                                if (len(self.paragraphs[-1][-1][0]) > 0
                                                and self.paragraphs[-1][-1][0][-1].tag == '{{{0}}}text'.format(nsmap['toraman'])):
                                                    self.paragraphs[-1][-1][0][-1].text += tb_paragraph_child.tail
                                                else:
                                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                                    self.paragraphs[-1][-1][0][-1].text = tb_paragraph_child.tail
                                            else:
                                                self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                                self.paragraphs[-1][-1][0][-1].text = tb_paragraph_child.tail

                                            tb_paragraph_child.tail = None
                                else:
                                    paragraph_continues = False

                            elif paragraph_child[0].tag == '{{{0}}}image'.format(self.nsmap['draw']):
                                if (paragraph_child.tail is not None
                                or paragraph_element.find('toraman:paragraph', nsmap) is not None
                                or paragraph_element.find('text:a', self.nsmap) is not None
                                or paragraph_element.find('text:span', self.nsmap) is not None
                                or paragraph_element.find('text:s', self.nsmap) is not None):
                                    image_copy = paragraph_child.__deepcopy__(True)
                                    image_copy.tail = None
                                    self.images.append(etree.tostring(image_copy))

                                    if paragraph_continues:
                                        if len(self.paragraphs[-1][-1]) != 1:
                                            self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])

                                        paragraph_element.remove(paragraph_child)

                                    else:
                                        self.paragraphs.append([[etree.Element('{{{0}}}run'.format(nsmap['toraman']))]])

                                        paragraph_element.replace(paragraph_child, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                                                no=str(len(self.paragraphs))))

                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}image'.format(self.t_nsmap['toraman']),
                                                                    no=str(len(self.images)),
                                                                    nsmap=self.t_nsmap))

                                    paragraph_continues = True

                        elif paragraph_child.tag == '{{{0}}}a'.format(self.nsmap['text']):
                            hyperlink_tag = paragraph_child.__deepcopy__(True)
                            hyperlink_tag.text = None
                            for hyperlink_child in hyperlink_tag:
                                hyperlink_tag.remove(hyperlink_child)
                            hyperlink_tag.tail = None
                            hyperlink_tag = etree.tostring(hyperlink_tag)

                            if not paragraph_continues:
                                self.paragraphs.append([[None, None, hyperlink_tag, 'beginning']])

                                paragraph_element.replace(paragraph_child, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                                        no=str(len(self.paragraphs))))

                                paragraph_continues = True

                            else:
                                self.paragraphs[-1].append([None, None, hyperlink_tag, 'beginning'])

                                paragraph_element.remove(paragraph_child)

                            if paragraph_child.text is not None:
                                self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                self.paragraphs[-1][-1][0][-1].text = paragraph_child.text

                            for a_element_child in paragraph_child:
                                if a_element_child.tag == '{{{0}}}span'.format(self.nsmap['text']):
                                    a_element_run_properties = a_element_child.attrib['{{{0}}}style-name'.format(self.nsmap['text'])]
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman'])), a_element_run_properties])
                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                    self.paragraphs[-1][-1][0][-1].text = a_element_child.text

                                if a_element_child.tail is not None:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                    self.paragraphs[-1][-1][0][-1].text = a_element_child.tail

                            self.paragraphs[-1].append([None, None, hyperlink_tag, 'end'])

                        elif paragraph_child.tag == '{{{0}}}span'.format(self.nsmap['text']):
                            extract_span(paragraph_child, paragraph_element, paragraph_continues)

                            paragraph_continues = True

                        elif paragraph_child.tag == '{{{0}}}line-break'.format(self.nsmap['text']):
                            if paragraph_continues:
                                if len(self.paragraphs[-1][-1]) != 1:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])

                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}br'.format(nsmap['toraman'])))

                                paragraph_element.remove(paragraph_child)

                        elif paragraph_child.tag == '{{{0}}}tab'.format(self.nsmap['text']):
                            if paragraph_continues:
                                if len(self.paragraphs[-1][-1]) != 1:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])

                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}tab'.format(nsmap['toraman'])))

                                paragraph_element.remove(paragraph_child)

                        elif paragraph_child.tag == '{{{0}}}s'.format(self.nsmap['text']):
                            if paragraph_continues:
                                if len(self.paragraphs[-1][-1]) == 1:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                    if (len(self.paragraphs[-1][-1][0]) > 0
                                    and self.paragraphs[-1][-1][0][-1].tag == '{{{0}}}text'.format(nsmap['toraman'])):
                                        self.paragraphs[-1][-1][0][-1].text += ' '
                                    else:
                                        self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                        self.paragraphs[-1][-1][0][-1].text = ' '
                                else:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                    self.paragraphs[-1][-1][0][-1].text = ' '

                                paragraph_element.remove(paragraph_child)

                        elif paragraph_child.tag == '{{{0}}}paragraph'.format(nsmap['toraman']):
                            pass
                        else:
                            if paragraph_continues:
                                self.miscellaneous_tags.append(etree.tostring(paragraph_child))
                                if len(self.paragraphs[-1][-1]) != 1:
                                    self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])

                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}{1}'.format(nsmap['toraman'], paragraph_child.tag.split('}')[1]),
                                                                                no=str(len(self.miscellaneous_tags))))

                                paragraph_element.remove(paragraph_child)

                        if paragraph_child.tail is not None:
                            if not paragraph_continues:
                                self.paragraphs.append([[etree.Element('{{{0}}}run'.format(nsmap['toraman']))]])

                                paragraph_element.insert(paragraph_element.index(paragraph_child) + 1, etree.Element('{{{0}}}paragraph'.format(nsmap['toraman']),
                                                                                                                    no=str(len(self.paragraphs))))

                                paragraph_continues = True

                            if len(self.paragraphs[-1][-1]) == 1:
                                if (len(self.paragraphs[-1][-1][0]) > 0
                                and self.paragraphs[-1][-1][0][-1].tag == '{{{0}}}text'.format(nsmap['toraman'])):
                                    self.paragraphs[-1][-1][0][-1].text += paragraph_child.tail
                                else:
                                    self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                    self.paragraphs[-1][-1][0][-1].text = paragraph_child.tail
                            else:
                                self.paragraphs[-1].append([etree.Element('{{{0}}}run'.format(nsmap['toraman']))])
                                self.paragraphs[-1][-1][0].append(etree.Element('{{{0}}}text'.format(nsmap['toraman'])))
                                self.paragraphs[-1][-1][0][-1].text = paragraph_child.tail

                            paragraph_child.tail = None

                    if (len(self.paragraphs[-1]) == 1
                    and len(self.paragraphs[-1][0]) == 1
                    and len(self.paragraphs[-1][0][0]) == 0):
                        self.paragraphs = self.paragraphs[:-1]

        # Filetype-specific processing ends here.

        toraman_link_template = etree.Element('{{{0}}}link'.format(self.t_nsmap['toraman']))

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
                        organised_paragraph.append(toraman_link)

                    for run_element in run[0]:
                        organised_paragraph.append(run_element)

                    if run[2] is not None:
                        toraman_link = etree.fromstring(toraman_link_template)
                        toraman_link.attrib['type'] = 'end'
                        organised_paragraph.append(toraman_link)

                    if run[1] is not None:
                        toraman_tag = etree.fromstring(toraman_tag_template)
                        toraman_tag.attrib['type'] = 'end'
                        organised_paragraph.append(toraman_tag)

                elif len(run) == 4:
                    if run[2] not in self.hyperlinks:
                            self.hyperlinks.append(run[2])

                    toraman_link = toraman_link_template.__deepcopy__(True)
                    toraman_link.attrib['no'] = str(self.hyperlinks.index(run[2])+1)
                    toraman_link.attrib['type'] = run[3]
                    organised_paragraph.append(toraman_link)

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
                    if toraman_element.text is None:
                        continue
                    _text = toraman_element.text.split(placeholders[1])
                    if _text[-1] is '':
                        _text = _text[:-1]
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

            for organised_segment in organised_paragraph[0]:
                if organised_segment.tag == '{{{0}}}non-text-segment'.format(self.t_nsmap['toraman']):
                    continue
                if (organised_segment[0].tag == '{{{0}}}tag'.format(self.t_nsmap['toraman'])
                and organised_segment[0].attrib['type'] == 'end'):
                    organised_segment.remove(organised_segment[0])
                if (organised_segment[-1].tag == '{{{0}}}tag'.format(self.t_nsmap['toraman'])
                and organised_segment[-1].attrib['type'] == 'beginning'):
                    organised_segment.remove(organised_segment[-1])

                active_ftags = []
                no_text_yet = True
                items_to_nontext = []
                perform_move = False
                for organised_segment_child in organised_segment:
                    if organised_segment_child.tag == '{{{0}}}tag'.format(self.t_nsmap['toraman']):
                        if organised_segment_child.attrib['type'] == 'beginning':
                            active_ftags.append(organised_segment_child.attrib['no'])
                        elif organised_segment_child.attrib['type'] == 'end':
                            if organised_segment_child.attrib['no'] in active_ftags:
                                active_ftags.remove(organised_segment_child.attrib['no'])
                            else:
                                _tag = organised_segment_child.__deepcopy__(True)
                                _tag.attrib['type'] = 'beginning'
                                organised_segment.insert(0, _tag)

                    if no_text_yet:
                        if organised_segment_child.tag == '{{{0}}}text'.format(self.t_nsmap['toraman']):
                            leading_space = regex.match(r'\s+', organised_segment_child.text)
                            if leading_space:
                                leading_space = leading_space.group()
                                if leading_space == organised_segment_child.text:
                                    items_to_nontext.append(organised_segment_child)
                                else:
                                    organised_segment_child.text = organised_segment_child.text[len(leading_space):]
                                    items_to_nontext.append(etree.Element('{{{0}}}text'.format(self.t_nsmap['toraman'])))
                                    items_to_nontext[-1].text = leading_space
                                perform_move = True
                            no_text_yet = False
                        elif organised_segment_child.tag == '{{{0}}}tag'.format(self.t_nsmap['toraman']):
                            items_to_nontext.append(organised_segment_child)

                        elif (organised_segment_child.tag == '{{{0}}}br'.format(self.t_nsmap['toraman'])
                        or organised_segment_child.tag == '{{{0}}}tab'.format(self.t_nsmap['toraman'])):
                            items_to_nontext.append(organised_segment_child)
                            perform_move = True
                        else:
                            no_text_yet = False

                else:
                    if active_ftags:
                        for active_ftag in active_ftags:
                            organised_segment.append(etree.Element('{{{0}}}tag'.format(self.t_nsmap['toraman'])))
                            organised_segment[-1].attrib['no'] = active_ftag
                            organised_segment[-1].attrib['type'] = 'end'
                        else:
                            active_ftags = []

                    if items_to_nontext and perform_move:
                        organised_segment_i = organised_paragraph[0].index(organised_segment)
                        if organised_paragraph[0][organised_segment_i-1].tag != '{{{0}}}non-text-segment'.format(self.t_nsmap['toraman']):
                            organised_paragraph[0].insert(organised_segment_i, etree.Element('{{{0}}}non-text-segment'.format(self.t_nsmap['toraman'])))
                        else:
                            organised_segment_i -= 1

                        for item_to_nontext in items_to_nontext:
                            if item_to_nontext.tag == '{{{0}}}.tag'.format(self.t_nsmap['toraman']):
                                if item_to_nontext.attrib['type'] == 'beginning':
                                    active_ftags.append(item_to_nontext.attrib['no'])
                                else:
                                    active_ftags.remove(item_to_nontext.attrib['no'])
                            organised_paragraph[0][organised_segment_i].append(item_to_nontext)
                        else:
                            if active_ftags:
                                _first_child = organised_paragraph[0][organised_segment_i+1][0]
                                if (_first_child.tag == '{{{0}}}.tag'.format(self.t_nsmap['toraman'])
                                and _first_child.attrib['type'] == 'end'
                                and _first_child.attrib['no'] in active_ftags):
                                    organised_paragraph[0][organised_segment_i].append(_first_child)
                                else:
                                    _tag = etree.Element('{{{0}}}.tag'.format(self.t_nsmap['toraman']))
                                    _tag.attrib['no'] = active_ftags[0]
                                    _tag.attrib['type'] = 'end'
                                    organised_paragraph[0][organised_segment_i].append(_tag)

                                    _tag = _tag.__deepcopy__(True)
                                    _tag.attrib['type'] = 'beginning'
                                    organised_paragraph[0][organised_segment_i+1].insert(0, _tag)

                    elif no_text_yet:
                        organised_segment.tag = '{{{0}}}non-text-segment'.format(self.t_nsmap['toraman'])

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

        if self.miscellaneous_tags:
            bilingual_file.append(etree.Element('{{{0}}}miscellaneous_tags'.format(self.t_nsmap['toraman'])))
            for mt_i in range(len(self.miscellaneous_tags)):
                new_t_element = etree.Element('{{{0}}}miscellaneous_tag'.format(self.t_nsmap['toraman']), no=str(mt_i + 1))
                new_t_element.text = self.miscellaneous_tags[mt_i]
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
