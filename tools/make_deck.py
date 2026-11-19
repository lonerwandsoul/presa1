"""Build the editable deck and a portable PDF (python-pptx, Pillow, DejaVu fonts)."""
from pathlib import Path
from io import BytesIO
import zipfile, xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR

ROOT=Path(__file__).resolve().parents[1]; assets=ROOT/'pres'
r=Presentation(); r.slide_width=Inches(13+1/3); r.slide_height=Inches(7.5)
cream='F8F1E2'; navy='152E43'; red='B82E2A'; peach='F9B49B'; white='FFFFFF'; gray='B8CBCF'
fonts=Path('/usr/share/fonts/truetype/dejavu'); pages=[]; k=1.5
def px(x): return Inches(x/96)
def rgb(c): return RGBColor.from_string(c)
def rect(x,y,w,h,c):
    a=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,px(x),px(y),px(w),px(h)); a.fill.solid(); a.fill.fore_color.rgb=rgb(c); a.line.fill.background(); draw.rectangle(tuple(round(i*k) for i in (x,y,x+w,y+h)),fill='#'+c)
def txt(x,y,t,size,c=navy,b=False,mono=False,serif=False):
    family='DejaVu '+('Sans Mono' if mono else 'Serif' if serif else 'Sans'); f=f"{family.replace(' ','')}{'-Bold' if b else ''}.ttf"; font=ImageFont.truetype(str(fonts/f),round(size*k)); width=draw.textlength(t,font=font)/k
    assert x+width < 1278,(t,x,width)
    a=s.shapes.add_textbox(px(x),px(y),px(width+15),px(size*1.55)); fr=a.text_frame; fr.word_wrap=False; fr.margin_left=fr.margin_top=fr.margin_bottom=fr.margin_right=0; fr.vertical_anchor=MSO_ANCHOR.TOP; p=fr.paragraphs[0]; p.space_before=Pt(0);p.space_after=Pt(0); run=p.add_run();run.text=t; run.font.name=family;run.font.size=Pt(size*.75);run.font.bold=b;run.font.color.rgb=rgb(c);draw.text((x*k,y*k),t,font=font,fill='#'+c)
def lines(x,y,ss,size,step,**kw):
    for i,t in enumerate(ss): txt(x,y+i*step,t,size,**kw)
def photo(name,x,y,w,h,center=(.5,.5)):
    im=ImageOps.fit(Image.open(assets/(name+'.jpeg')).convert('RGB'),(round(w*k),round(h*k)),method=Image.Resampling.LANCZOS,centering=center);buf=BytesIO();im.save(buf,'JPEG',quality=94);buf.seek(0);s.shapes.add_picture(buf,px(x),px(y),px(w),px(h)); page.paste(im,(round(x*k),round(y*k)))
# Speaker notes follow the supplied scenario, including discussion prompts and reference timings.
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with zipfile.ZipFile(next(assets.glob('*.docx'))) as z: doc=ET.fromstring(z.read('word/document.xml'))
notes={}; n=0
for p in doc.findall('.//w:p',ns):
    t=''.join(e.text or '' for e in p.findall('.//w:t',ns))
    if t.startswith('Слайд '): n=int(t.split()[1].rstrip('.')); notes[n]=[]
    elif t=='ПОРЯДОК ДЕЙСТВИЙ':break
    elif n and t.startswith(('Сказать:', 'Видео №','Вопрос аудитории:', 'Рефлексия:', 'Заключительный вопрос:', 'Выслушать')):notes[n].append(t.replace('Видео №','Кадр из эпизода №'))
def start(bg,label):
    global s,page,draw
    s=r.slides.add_slide(r.slide_layouts[6]); s.background.fill.solid();s.background.fill.fore_color.rgb=rgb(bg);page=Image.new('RGB',(1920,1080),'#'+bg);pages.append(page);draw=ImageDraw.Draw(page);idx=len(pages);s.notes_slide.notes_text_frame.text='\n\n'.join(notes[idx]); c=navy if bg==cream else cream
    txt(48,25,'КИНОКЛУБ   /   1+1',18,c,True,True);txt(442,25,label,18,c,mono=True);txt(1130,25,f'{idx:02} / 12',18,c,mono=True);rect(48,60,1184,2,c)
def tag(x,y,t,bg=red,c=white):
    w=20+len(t)*10;rect(x,y,w,31,bg);txt(x+10,y+4,t,16,c,True,True)
def cap(x,y,t,c=navy):txt(x,y,t,16,c,True,True)
def foot(t,c=navy):rect(48,644,1184,2,c);txt(48,661,t,18,c,mono=True)
start(cream,'СПЕЦПОКАЗ     •     ФРАНЦИЯ, 2011');photo('sowmestkadr2',655,85,577,550,(.5,.42));tag(676,104,'INTOUCHABLES');cap(48,95,'ФИЛЬМ О ТОМ, КАК МЫ ОБЩАЕМСЯ');txt(28,117,'1+1',210,red,True);lines(50,375,['Как подружиться,','если у вас даже','плейлисты разные?'],41,51,b=True);foot('РЕЖИССЁРЫ   /   Оливье Накаш и Эрик Толедано')
start(cream,'В РОЛЯХ');txt(48,79,'Два человека, два плейлиста',54,b=True)
for x,name,title,actor,desc,badge in [(48,'philip','Филипп','Франсуа Клюзе',['Ценит живопись и классику.','Ищет помощника, а не сочувствия.'],'01  /  КЛАССИКА'),(656,'driss','Дрисс','Омар Си',['Помощник Филиппа. Говорит прямо,','любит фанк и танцует первым.'],'02  /  ФАНК')]:
    photo(name,x,171,576,285);tag(x+16,408,badge);txt(x,470,title,46,b=True);txt(x,533,actor,21,red,True,True);lines(x,595,desc,25,34)
start(navy,'ПРОЛОГ     /     ЗНАКОМСТВО');photo('sowmestkadr1',48,89,731,545,(.48,.5));tag(68,108,'ПАРИЖ');cap(813,111,'С ЧЕГО ВСЁ НАЧАЛОСЬ',peach);lines(810,181,['Вообще-то','это было','собеседование.'],41,51,c=cream,b=True);lines(813,393,['Дрисс приходит за подписью,','а получает работу. За шутками','и спорами о вкусах постепенно','возникает дружба.'],23,34,c=cream);foot('ДАЛЬШЕ В ПРОГРАММЕ   /   День рождения и один очень дорогой холст',cream)

def scene(label,num,title,body,im,time,caption,center=(.5,.5),bg=cream):
    start(bg,label);fg=navy if bg==cream else cream;accent=red if bg==cream else peach
    cap(48,112,'КАДР   /   '+num,accent);lines(46,185,title,51,63,c=fg,b=True);lines(49,423,body,24,35,c=fg);photo(im,562,91,670,488,center);rect(562,579,670,45,red);txt(577,589,'TC   '+time,18,cream,True,True);foot(caption,fg)
scene('ИСТОРИЯ 01     /     МУЗЫКА','01','Праздник\nпо нотам'.split('\n'),['В честь Филиппа играют классику.','Гости слушают. Дриссу не хватает','движения — но нужно ли оно всем?'],'довключениямузыки','01:09:00 — 01:09:25','ДЕНЬ РОЖДЕНИЯ   /   До смены музыки')
scene('ИСТОРИЯ 01     /     МУЗЫКА','02',['А если','потанцевать?'],['Дрисс ставит Boogie Wonderland','и первым выходит танцевать.','Гости подхватывают его движения.'],'послевключениямузыки','≈ 01:11:00 — 01:11:25','ДЕНЬ РОЖДЕНИЯ   /   После смены музыки',(.39,.5),navy)
start(cream,'ИСТОРИЯ 01     /     РАЗБОР');txt(48,84,'Почему гости включились?',53,b=True);cap(49,177,'ДВА СПОСОБА БЫТЬ НА ПРАЗДНИКЕ',red)
for x,img,a,b in [(48,'довключениямузыки','Слушать','У каждого свой музыкальный вкус.'),(656,'послевключениямузыки','Танцевать','Личный пример сработал лучше уговоров.')]:
    photo(img,x,226,576,207,(.5,.37));txt(x,448,a,42,b=True);txt(x,518,b,22)
rect(48,605,1184,70,red);txt(67,624,'РИСК   /   Сменить музыку без спроса — и перебить любимую мелодию другого.',21,cream,True)
start(red,'ИСТОРИЯ 01     /     НА ЗАМЕТКУ');lines(46,118,['Плейлист','для всех?'],76,89,c=cream,b=True);lines(50,374,['Пусть у гостя будет выбор:','подпевать, танцевать или','просто слушать.'],28,39,c=cream);rect(645,101,588,520,cream);cap(674,132,'ПАМЯТКА ОРГАНИЗАТОРУ');
for i,(a,b) in enumerate([('Спросить заранее','Что хотелось бы услышать гостям?'),('Чередовать форматы','После концерта предложить танцы.'),('Оставить выбор','Приглашать, но не тянуть за руку.')]):
    y=207+i*128;cap(675,y,f'0{i+1}',red);txt(728,y-9,a,30,b=True);txt(728,y+35,b,20);rect(675,y+82,522,1,navy)
foot('ПРИНЦИП   /   Хорошая программа оставляет место для разных гостей.',cream)
scene('ИСТОРИЯ 02     /     ЖИВОПИСЬ','03',['И это —','искусство?'],['В галерее Филипп задерживается','у абстрактной картины. Дрисса','скорее удивляет её цена.'],'первыйфрагментнепониманияискусства','≈ 00:31:20 — 00:31:45','ГАЛЕРЕЯ   /   Один холст, две реакции',(.5,.5))
scene('ИСТОРИЯ 02     /     ЖИВОПИСЬ','04',['От шутки','до кисти'],['Филипп делится своим интересом','к живописи без лекций и споров.','Дрисс пробует рисовать сам.'],'созданиекартины','≈ 00:59:57 — 01:00:20','МАСТЕРСКАЯ   /   Смотреть со стороны и делать самому — разный опыт',(.48,.5),navy)
start(cream,'ИСТОРИЯ 02     /     ПРИЗНАНИЕ');cap(48,104,'ЦЕНА КАРТИНЫ ДРИССА');txt(37,131,'11 000 €',145,red,True);photo('продажа картины',48,336,564,278);lines(654,344,['Так кисть окупилась.','А что изменилось?'],35,46,b=True);lines(655,477,['Филипп продал картину. Для Дрисса','сюрприз не только в сумме: кто-то','захотел повесить его работу у себя.'],24,34);foot('КАДР 05   /   ≈ 01:16:00 — 01:16:25')
start(cream,'ИСТОРИЯ 02     /     РАЗБОР');lines(48,87,['Как увлечь искусством','и не отбить интерес?'],49,60,b=True)
for x,bg,fg,label,head,tail in [(48,navy,cream,'01   /   РИСК','Читать нотации',['Если объяснять, что «надо» любить,','разговор о вкусах станет экзаменом.']),(656,red,cream,'02   /   АЛЬТЕРНАТИВА','Дать попробовать',['Спросить, что человек видит в картине.','Предложить кисть — не оценку.'])]:
    rect(x,268,576,285,bg);cap(x+25,293,label,fg);txt(x+24,365,head,36,fg,True);lines(x+25,447,tail,22,32,c=fg)
foot('РЕЗУЛЬТАТ   /   Дрисс уже не спорит о чужой картине — у него есть своя.')
start(navy,'ФИНАЛ     /     РЕФЛЕКСИЯ');photo('sowmestkadr3',792,91,440,417,(.5,.45));tag(809,109,'FIN');lines(48,101,['Что бы выбрали вы:','выйти на танцпол','или взять кисть?'],51,64,c=cream,b=True);lines(51,353,['Когда легче решиться: если кто-то начинает','первым или если можно попробовать самому?','Вспомните свой пример.'],24,36,c=cream);rect(48,556,1184,2,peach);cap(48,583,'ВЫВОД',peach);lines(215,573,['Дрисс пригласил гостей в танец, Филипп поддержал художника.','Оба оставили другому право решать, участвовать ли.'],25,37,c=cream)
name=ROOT/'1plus1';r.core_properties.title='1+1. Как подружиться, если у вас даже плейлисты разные?';r.save(str(name)+'.pptx');pages[0].save(str(name)+'.pdf',save_all=True,append_images=pages[1:],resolution=144,quality=92)
