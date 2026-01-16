"""
ICS (iCalendar) dosyalarını parse eden modül
"""
from icalendar import Calendar
from datetime import datetime
from typing import List, Dict
import os
import re
from html.parser import HTMLParser


class HTMLTextExtractor(HTMLParser):
    """HTML'den metin çıkaran ve formatlamayı koruyan parser"""
    
    def __init__(self):
        super().__init__()
        self.text_parts = []
        self.bold_stack = []
        self.italic_stack = []
    
    def handle_starttag(self, tag, attrs):
        tag_lower = tag.lower()
        if tag_lower in ['b', 'strong']:
            self.bold_stack.append(True)
        elif tag_lower in ['i', 'em']:
            self.italic_stack.append(True)
        elif tag_lower == 'br':
            self.text_parts.append({'text': '\n', 'bold': False, 'italic': False})
    
    def handle_endtag(self, tag):
        tag_lower = tag.lower()
        if tag_lower in ['b', 'strong']:
            if self.bold_stack:
                self.bold_stack.pop()
        elif tag_lower in ['i', 'em']:
            if self.italic_stack:
                self.italic_stack.pop()
    
    def handle_data(self, data):
        bold = len(self.bold_stack) > 0
        italic = len(self.italic_stack) > 0
        if data.strip():  # Boş olmayan metinleri ekle
            self.text_parts.append({'text': data, 'bold': bold, 'italic': italic})
    
    def get_formatted_text(self):
        """Formatlanmış metin parçalarını döner"""
        return self.text_parts
    
    def get_plain_text(self):
        """Sadece metni döner (formatlama olmadan)"""
        return ''.join(part['text'] for part in self.text_parts)


def clean_html(html_text: str) -> str:
    """
    HTML etiketlerini temizler ve düz metne dönüştürür
    
    Args:
        html_text: HTML içeren metin
        
    Returns:
        Temizlenmiş metin
    """
    if not html_text:
        return ''
    
    # HTML entity'leri decode et
    html_text = html_text.replace('&nbsp;', ' ')
    html_text = html_text.replace('&amp;', '&')
    html_text = html_text.replace('&lt;', '<')
    html_text = html_text.replace('&gt;', '>')
    html_text = html_text.replace('&quot;', '"')
    html_text = html_text.replace('&#39;', "'")
    
    # <br>, <BR>, <br/> gibi etiketleri yeni satıra çevir
    html_text = re.sub(r'<br\s*/?>', '\n', html_text, flags=re.IGNORECASE)
    
    # Diğer HTML etiketlerini kaldır
    html_text = re.sub(r'<[^>]+>', '', html_text)
    
    # Çoklu boşlukları tek boşluğa çevir
    html_text = re.sub(r'\s+', ' ', html_text)
    
    # Başta ve sonda boşlukları temizle
    html_text = html_text.strip()
    
    return html_text


def parse_html_with_formatting(html_text: str) -> List[Dict]:
    """
    HTML'den formatlanmış metin parçalarını çıkarır
    
    Args:
        html_text: HTML içeren metin
        
    Returns:
        Formatlanmış metin parçaları listesi
    """
    if not html_text:
        return [{'text': '', 'bold': False, 'italic': False}]
    
    # HTML entity'leri decode et
    html_text = html_text.replace('&nbsp;', ' ')
    html_text = html_text.replace('&amp;', '&')
    html_text = html_text.replace('&lt;', '<')
    html_text = html_text.replace('&gt;', '>')
    html_text = html_text.replace('&quot;', '"')
    html_text = html_text.replace('&#39;', "'")
    
    # HTML parser ile formatlanmış metni çıkar
    parser = HTMLTextExtractor()
    parser.feed(html_text)
    
    parts = parser.get_formatted_text()
    return parts if parts else [{'text': '', 'bold': False, 'italic': False}]


class ICSParser:
    """ICS dosyalarını okuyup etkinlikleri çıkaran sınıf"""
    
    def __init__(self, ics_file_path: str):
        """
        Args:
            ics_file_path: ICS dosyasının yolu
        """
        self.ics_file_path = ics_file_path
        self.events = []
    
    def parse(self) -> List[Dict]:
        """
        ICS dosyasını parse eder ve etkinlik listesi döner
        
        Returns:
            Etkinlik listesi (her etkinlik dict formatında)
        """
        if not os.path.exists(self.ics_file_path):
            raise FileNotFoundError(f"ICS dosyası bulunamadı: {self.ics_file_path}")
        
        with open(self.ics_file_path, 'rb') as f:
            calendar_data = f.read()
        
        calendar = Calendar.from_ical(calendar_data)
        events = []
        
        for component in calendar.walk('VEVENT'):
            event = {}
            
            # Başlık (HTML temizle)
            summary = str(component.get('SUMMARY', ''))
            event['summary'] = clean_html(summary)
            event['summary_formatted'] = parse_html_with_formatting(summary)
            
            # Başlangıç tarihi
            dtstart = component.get('DTSTART')
            if dtstart:
                dt = dtstart.dt
                if isinstance(dt, datetime):
                    event['start'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    event['start'] = str(dt)
            else:
                event['start'] = ''
            
            # Bitiş tarihi
            dtend = component.get('DTEND')
            if dtend:
                dt = dtend.dt
                if isinstance(dt, datetime):
                    event['end'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    event['end'] = str(dt)
            else:
                event['end'] = ''
            
            # Açıklama (HTML temizle)
            description = str(component.get('DESCRIPTION', ''))
            event['description'] = clean_html(description)
            event['description_formatted'] = parse_html_with_formatting(description)
            
            # Konum
            event['location'] = str(component.get('LOCATION', ''))
            
            # Organizer
            organizer = component.get('ORGANIZER')
            if organizer:
                event['organizer'] = str(organizer)
            else:
                event['organizer'] = ''
            
            # URL
            url = component.get('URL')
            if url:
                event['url'] = str(url)
            else:
                event['url'] = ''
            
            # UID
            event['uid'] = str(component.get('UID', ''))
            
            events.append(event)
        
        self.events = events
        return events
    
    def get_events_count(self) -> int:
        """Parse edilen etkinlik sayısını döner"""
        return len(self.events)
