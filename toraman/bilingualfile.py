import os

from lxml import etree

from .utils import get_current_time_in_utc


class BilingualFile:
    def __init__(self, file_path):
        self.hyperlinks = []
        self.images = []
        self.paragraphs = []
        self.tags = []
        self.miscellaneous_tags = []

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

        if self.xml_root.find('toraman:miscellaneous_tags', self.t_nsmap) is not None:
            for xml_mtag in self.xml_root.find('toraman:miscellaneous_tags', self.t_nsmap):
                self.miscellaneous_tags.append(xml_mtag.text)

        if self.xml_root.find('toraman:hyperlinks', self.t_nsmap) is not None:
            for xml_hl in self.xml_root.find('toraman:hyperlinks', self.t_nsmap):
                self.hyperlinks.append(xml_hl.text)

        if self.xml_root.find('toraman:images', self.t_nsmap) is not None:
            for xml_image in self.xml_root.find('toraman:images', self.t_nsmap):
                self.images.append(xml_image.text)

    def generate_target_translation(self, source_file_path, output_directory):
        from hashlib import sha256
        import zipfile

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

        final_paragraphs = []

        if self.file_type == 'docx':

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

        elif self.file_type == 'odt':

            for target_paragraph in target_paragraphs:
                active_ftags = []
                active_links = []
                final_paragraph = []

                for child in target_paragraph:
                    if child.tag == '{{{0}}}text'.format(self.nsmap['toraman']):
                        if child.text is None:
                            continue
                        if active_links and active_ftags:
                            if len(final_paragraph[-1][-1]) == 0:
                                if final_paragraph[-1][-1].text is None:
                                    final_paragraph[-1][-1].text = ''
                                final_paragraph[-1][-1].text += child.text
                            else:
                                if final_paragraph[-1][-1][-1].tail is None:
                                    final_paragraph[-1][-1][-1].tail = ''
                                final_paragraph[-1][-1][-1].tail += child.text
                        elif active_links or active_ftags:
                            if len(final_paragraph[-1]) == 0:
                                if final_paragraph[-1].text is None:
                                    final_paragraph[-1].text = ''
                                final_paragraph[-1].text += child.text
                            else:
                                if final_paragraph[-1][-1].tail is None:
                                    final_paragraph[-1][-1].tail = ''
                                final_paragraph[-1][-1].tail += child.text
                        else:
                            final_paragraph.append(child.text)

                    elif child.tag == '{{{0}}}tag'.format(self.nsmap['toraman']):
                        if child.attrib['type'] == 'beginning':
                            if child.attrib['no'] not in active_ftags:
                                active_ftags.append(child.attrib['no'])

                        else:
                            if child.attrib['no'] in active_ftags:
                                active_ftags.remove(child.attrib['no'])

                        if active_ftags and active_links:
                            final_paragraph[-1].append(etree.Element('{{{0}}}span'.format(self.nsmap['text'])))
                            final_paragraph[-1][-1].attrib['{{{0}}}style-name'.format(self.nsmap['text'])] = self.tags[int(active_ftags[0])-1]

                        elif active_ftags:
                            final_paragraph.append(etree.Element('{{{0}}}span'.format(self.nsmap['text'])))
                            final_paragraph[-1].attrib['{{{0}}}style-name'.format(self.nsmap['text'])] = self.tags[int(active_ftags[0])-1]

                    elif child.tag == '{{{0}}}image'.format(self.nsmap['toraman']):
                        if active_ftags and active_links:
                            final_paragraph[-1][-1].append(etree.fromstring(self.images[int(child.attrib['no'])-1]))
                        elif active_ftags or active_links:
                            final_paragraph[-1].append(etree.fromstring(self.images[int(child.attrib['no'])-1]))
                        else:
                            final_paragraph.append(etree.fromstring(self.images[int(child.attrib['no'])-1]))

                    elif child.tag == '{{{0}}}link'.format(self.nsmap['toraman']):
                        if child.attrib['type'] == 'beginning':
                            if child.attrib['no'] not in active_links:
                                active_links.append(child.attrib['no'])
                        else:
                            if child.attrib['no'] in active_links:
                                active_links.remove(child.attrib['no'])

                        if active_links:
                            final_paragraph.append(etree.fromstring(self.hyperlinks[int(child.attrib['no'])-1]))
                            if active_ftags:
                                final_paragraph[-1].append(etree.Element('{{{0}}}span'.format(self.nsmap['text'])))
                                final_paragraph[-1][-1].attrib['{{{0}}}style-name'.format(self.nsmap['text'])] = self.tags[int(active_ftags[0])-1]

                    elif child.tag == '{{{0}}}br'.format(self.nsmap['toraman']):
                        if active_ftags and active_links:
                            final_paragraph[-1][-1].append(etree.Element('{{{0}}}line-break'.format(self.nsmap['text'])))
                        elif active_ftags or active_links:
                            final_paragraph[-1].append(etree.Element('{{{0}}}line-break'.format(self.nsmap['text'])))
                        else:
                            final_paragraph.append(etree.Element('{{{0}}}line-break'.format(self.nsmap['text'])))

                    elif child.tag == '{{{0}}}tab'.format(self.nsmap['toraman']):
                        if active_ftags and active_links:
                            final_paragraph[-1][-1].append(etree.Element('{{{0}}}tab'.format(self.nsmap['text'])))
                        elif active_ftags or active_links:
                            final_paragraph[-1].append(etree.Element('{{{0}}}tab'.format(self.nsmap['text'])))
                        else:
                            final_paragraph.append(etree.Element('{{{0}}}tab'.format(self.nsmap['text'])))

                    else:
                        if active_ftags and active_links:
                            final_paragraph[-1][-1].append(etree.fromstring(self.miscellaneous_tags[int(child.attrib['no'])-1]))
                        elif active_ftags or active_links:
                            final_paragraph[-1].append(etree.fromstring(self.miscellaneous_tags[int(child.attrib['no'])-1]))
                        else:
                            final_paragraph.append(etree.fromstring(self.miscellaneous_tags[int(child.attrib['no'])-1]))


                final_paragraphs.append(final_paragraph)

            for internal_file in self.xml_root[-1]:
                internal_file = internal_file[0]
                for paragraph_placeholder in internal_file.findall('.//toraman:paragraph', self.t_nsmap):
                    paragraph_placeholder_parent = paragraph_placeholder.getparent()
                    placeholder_i = paragraph_placeholder_parent.index(paragraph_placeholder)
                    child_i = placeholder_i

                    for final_paragraph_child in final_paragraphs[int(paragraph_placeholder.attrib['no'])-1]:
                        if type(final_paragraph_child) == str:
                            if child_i == 0:
                                if paragraph_placeholder_parent.text is None:
                                    paragraph_placeholder_parent.text = ''
                                paragraph_placeholder_parent.text += final_paragraph_child
                            else:
                                if paragraph_placeholder_parent[child_i-1].tail is None:
                                    paragraph_placeholder_parent[child_i-1].tail = ''
                                paragraph_placeholder_parent[child_i-1].tail += final_paragraph_child
                        else:
                            if child_i == placeholder_i:
                                paragraph_placeholder_parent.replace(paragraph_placeholder, final_paragraph_child)
                            else:
                                paragraph_placeholder_parent.insert(child_i, final_paragraph_child)

                            child_i += 1

                    else:
                        if child_i == 0 and paragraph_placeholder in paragraph_placeholder_parent:
                            paragraph_placeholder_parent.remove(paragraph_placeholder)

        # Filetype-specific processing ends here.

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

    def update_segment(self, segment_status, segment_target, paragraph_no, segment_no, author_id):

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

        if 'creationdate' in xml_segment.attrib:
            xml_segment.attrib['changedate'] = get_current_time_in_utc()
            xml_segment.attrib['changeid'] = author_id
        else:
            xml_segment.attrib['creationdate'] = get_current_time_in_utc()
            xml_segment.attrib['creationid'] = author_id

        self.paragraphs[paragraph_no - 1][sub_p_id] = segment
