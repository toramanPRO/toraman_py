import os

import Levenshtein
from lxml import etree

from .utils import get_current_time_in_utc, nsmap, segment_to_tm_segment
from .version import __version__


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
            self.translation_memory[0].attrib['creationdate'] = get_current_time_in_utc()
            self.translation_memory[0].attrib['datatype'] = 'PlainText'
            self.translation_memory[0].attrib['segtype'] = 'sentence'
            self.translation_memory[0].attrib['adminlang'] = 'en'
            self.translation_memory[0].attrib['srclang'] = self.source_language
            self.translation_memory[0].attrib['trgtlang'] = self.target_language
            self.translation_memory.append(etree.Element('{{{0}}}body'.format(nsmap['toraman'])))

            self.translation_memory.getroottree().write(self.tm_path,
                                           encoding='UTF-8',
                                           xml_declaration=True)

    def lookup(self, source_segment, match=0.7, convert_segment=True):
        _segment_hits = []

        if convert_segment:
            segment_query = segment_to_tm_segment(source_segment)
        else:
            segment_query = source_segment

        for translation_unit in self.translation_memory[1]:
            levenshtein_ratio = Levenshtein.ratio(translation_unit[1].text, segment_query)
            if levenshtein_ratio >= match:
                _segment_hits.append((levenshtein_ratio,
                                    translation_unit[0].__deepcopy__(True),
                                    translation_unit[2].__deepcopy__(True)))
        else:
            _segment_hits.sort(reverse=True)

        return _segment_hits

    def submit_segment(self, source_segment, target_segment, author_id):
        segment_query = segment_to_tm_segment(source_segment)

        for translation_unit in self.translation_memory[1]:
            if translation_unit[1].text == segment_query:
                translation_unit[2] = target_segment.__deepcopy__(True)
                translation_unit.attrib['changedate'] = get_current_time_in_utc()
                translation_unit.attrib['changeid'] = author_id
                break
        else:
            translation_unit = etree.Element('{{{0}}}tu'.format(nsmap['toraman']))
            translation_unit.attrib['creationdate'] = get_current_time_in_utc()
            translation_unit.attrib['creationid'] = author_id
            translation_unit.append(source_segment.__deepcopy__(True))
            translation_unit.append(etree.Element('{{{0}}}query'.format(nsmap['toraman'])))
            translation_unit[1].text = segment_query
            translation_unit.append(target_segment.__deepcopy__(True))

            self.translation_memory[1].append(translation_unit)

        self.translation_memory.getroottree().write(self.tm_path,
                                                    encoding='UTF-8',
                                                    xml_declaration=True)
