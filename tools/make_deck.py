"""Generate the 1+1 pop-zine slides (Pillow, python-pptx, DejaVu)."""
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageFilter, ImageChops
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pathlib import Path
from io import BytesIO
from functools import lru_cache
import math, random, zipfile, xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[1]; assets=ROOT/'pres'; K=1.5; W,H=1280,720
ink='#1C1425';milk='#FFF5E6';pink='#FF86C8';hot='#FF408F';lime='#D6FF48';blue='#4138D3';lilac='#C2B3FF';orange='#FF8353'; mint='#A0EECC'
R=Presentation();R.slide_width=Inches(13+1/3);R.slide_height=Inches(7.5);pages=[]
# Speaker notes follow the supplied scenario, including discussion prompts and reference timings.
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
with zipfile.ZipFile(next(assets.glob('*.docx'))) as z: doc=ET.fromstring(z.read('word/document.xml'))
notes={}; n=0
for p in doc.findall('.//w:p',ns):
    t=''.join(e.text or '' for e in p.findall('.//w:t',ns))
    if t.startswith('Слайд '): n=int(t.split()[1].rstrip('.')); notes[n]=[]
    elif t=='ПОРЯДОК ДЕЙСТВИЙ':break
    elif n and t.startswith(('Сказать:', 'Видео №','Вопрос аудитории:', 'Рефлексия:', 'Заключительный вопрос:', 'Выслушать')):notes[n].append(t.replace('Видео №','Кадр из эпизода №'))
def sc(*a):return tuple(round(v*K) for v in a)
@lru_cache(None)
def font(sz,typ):
    f=ImageFont.truetype(str(ROOT/'tools/fonts/Unbounded.ttf' if typ=='h' else Path('/usr/share/fonts/truetype/dejavu')/('DejaVuSansMono-Bold.ttf' if typ=='m' else 'DejaVuSans.ttf')),round(sz*K));
    if typ=='h':f.set_variation_by_axes([830]);
    return f
def rect(x,y,w,h,c):d.rectangle(sc(x,y,x+w,y+h),c)
def ell(x,y,w,h,c,outline=None,lw=2):d.ellipse(sc(x,y,x+w,y+h),c,outline,round(lw*K))
def line(coords,c,w=2):d.line([sc(*p) for p in coords],c,round(w*K),joint='curve')
def poly(pts,c):d.polygon([sc(*p) for p in pts],c)
def text(x,y,t,sz,c=ink,typ='h',mw=None):
    while mw and font(sz,typ).getlength(t)>mw*K:sz-=1
    d.text(sc(x,y),t,font=font(sz,typ),fill=c,anchor='lt');return font(sz,typ).getlength(t)/K
def title(x,y,ss,sz,step,mw,c=ink):
    sz=min([sz]+[max(8,math.floor(sz*mw*K/font(sz,'h').getlength(a))) for a in ss]);
    for i,t in enumerate(ss):text(x,y+i*step,t,sz,c,mw=mw)
def body(x,y,ss,sz=24,c=ink,step=None):
    # The captions remain editable PowerPoint text; headings use the bundled display face.
    for i,t in enumerate(ss):captions.append((x,y+i*(step or sz*1.42),t,sz,c))
def sticker(x,y,t,bg,angle=0,sz=17,fg=ink):
    f=font(sz,'m'); w=round(f.getlength(t)+34*K); h=round((sz+20)*K);a=Image.new('RGBA',(w+7,h+7));g=ImageDraw.Draw(a);g.rounded_rectangle((5,5,w+5,h+5),radius=h//2,fill=ink);g.rounded_rectangle((0,0,w,h),radius=h//2,fill=bg,outline=ink,width=3);g.text((w/2,h/2),t,font=f,fill=fg,anchor='mm');a=a.rotate(angle,Image.Resampling.BICUBIC,expand=True);im.paste(a,sc(x,y),a)
def checker(x,y,n,m,cell,c1,c2):
    rect(x,y,n*cell,m*cell,c1)
    for j in range(m):
        for i in range((j+1)%2,n,2):rect(x+i*cell,y+j*cell,cell,cell,c2)
def star(x,y,r,c,angle=0):
    pts=[]
    for j in range(8):
        a=angle+j*math.pi/4;v=r if j%2==0 else r*.23;pts.append((x+math.cos(a)*v,y+math.sin(a)*v))
    poly(pts,c)
def record(x,y,r,label):
    ell(x-r,y-r,r*2,r*2,ink)
    for j in range(4,10):
        v=r*j/10;ell(x-v,y-v,v*2,v*2,None,'#72677E',1)
    v=r*.32;ell(x-v,y-v,v*2,v*2,label);ell(x-4,y-4,8,8,ink)
def pic(name,x,y,w,h,angle=0,ctr=(.5,.5),frame=True):
    a=ImageOps.fit(Image.open(assets/(name+'.jpeg')).convert('RGB'),sc(w,h),centering=ctr,method=Image.Resampling.LANCZOS).convert('RGBA');
    if frame:a=ImageOps.expand(ImageOps.expand(a,round(6*K),milk),round(2*K),ink)
    a=a.rotate(angle,Image.Resampling.BICUBIC,expand=True);sh=Image.new('RGBA',a.size,ink);im.paste(sh,sc(x+8,y+9),a.getchannel('A'));im.paste(a,sc(x,y),a)
def dots(x,y,w,h,c):
    for j in range(0,h,14):
        for i in range(0,w,14):ell(x+i+(j//14%2)*5,y+j,2.5,2.5,c)
def cutout(x,y,w):
    src=Image.open(assets/'sowmestkadr2.jpeg').convert('RGB');points=[(51,476),(52,432),(24,407),(10,386),(4,368),(6,338),(10,323),(15,320),(24,294),(39,260),(54,229),(66,210),(82,192),(96,178),(120,164),(151,149),(164,143),(184,123),(198,111),(202,88),(209,76),(222,65),(242,60),(257,62),(271,69),(281,81),(288,95),(289,111),(293,117),(294,133),(291,155),(298,168),(314,178),(333,190),(342,213),(348,246),(351,300),(353,319),(354,334),(346,338),(342,360),(336,379),(339,390),(352,402),(376,413),(406,428),(418,446),(428,464),(443,476)]
    mask=Image.new('L',src.size);ImageDraw.Draw(mask).polygon([(a*src.width/480,b*src.height/476) for a,b in points],255);src=ImageOps.colorize(ImageOps.autocontrast(ImageOps.grayscale(src)),ink,milk).convert('RGBA');src.putalpha(mask);src=src.resize(sc(w,w*476/480),Image.Resampling.LANCZOS);src=ImageOps.expand(src,12,(0,0,0,0));out=src.getchannel('A').filter(ImageFilter.MaxFilter(17));im.paste(milk,sc(x,y)+tuple(a+b for a,b in zip(sc(x,y),src.size)),out);im.paste(src,sc(x,y),src)
def begin(bg):
    global im,d,captions;im=Image.new('RGB',sc(W,H),bg);d=ImageDraw.Draw(im);captions=[]
def meta(t,fg=ink):
    text(42,28,'1+1  /  КИНОЗИН',16,fg,'m');text(435,28,t,16,fg,'m');text(1160,28,f'{len(pages)+1:02} / 12',16,fg,'m')
def tc(t,x,y):sticker(x,y,'КАДР  /  '+t,milk,0,15)
def end():
    # Keep the designed artwork identical on machines without the display font.
    s=R.slides.add_slide(R.slide_layouts[6]);buf=BytesIO();im.save(buf,'PNG',optimize=True);buf.seek(0);s.shapes.add_picture(buf,0,0,R.slide_width,R.slide_height)
    for x,y,t,sz,c in captions:
        w=font(sz,'b').getlength(t)/K; assert x+w<1278,(t,w)
        a=s.shapes.add_textbox(Inches(x/96),Inches((y-sz*.18)/96), Inches((w+34)/96), Inches(sz*1.6/96));fr=a.text_frame;fr.word_wrap=False;fr.margin_top=fr.margin_left=fr.margin_right=fr.margin_bottom=0;p=fr.paragraphs[0];run=p.add_run();run.text=t;run.font.size=Pt(sz*.75);run.font.name='DejaVu Sans';run.font.color.rgb=RGBColor.from_string(c[1:]);text(x,y,t,sz,c,'b')
    pages.append(im);s.notes_slide.notes_text_frame.text='\n\n'.join(notes[len(pages)])
# 01 / cover: gig poster, real die-cut portrait and optical type
begin(pink);ell(684,-95,752,752,lime);dots(900,0,390,720,ink);checker(0,655,66,4,20,ink,milk);meta('СПЕЦВЫПУСК   ✳   ФРАНЦИЯ / 2011');text(36,135,'1+1',228,ink,mw=720);sticker(48,75,'INTOUCHABLES  •  2011',milk,4);cutout(640,47,705);star(1164,147,65,hot,.2);star(724,158,41,blue,.15);title(48,406,['Дружба не по','плейлисту'],51,63,630);sticker(40,555,'КЛАССИКА × ФАНК × ЖИВОПИСЬ',lime,-3,19);rect(0,688,1280,32,ink);text(42,697,'РЕЖИССЁРЫ: ОЛИВЬЕ НАКАШ И ЭРИК ТОЛЕДАНО',15,milk,'m');end()
# 02 / matching collector cards
begin(lilac);meta('КАСТИНГ  /  КТО ЕСТЬ КТО');text(42,81,'Два разных вайба',63,mw=1050);star(1188,120,46,lime,.13)
for x,img,name,actor,tag,desc,c,ctr in [(48,'philip','Филипп','Франсуа Клюзе','01 / КЛАССИКА','Любит живопись и камерную музыку.',pink,(.55,.36)),(655,'driss','Дрисс','Омар Си','02 / ФАНК','Помощник Филиппа. Танцует первым.',lime,(.6,.4))]:
    rect(x+9,193,571,476,ink);rect(x,184,571,476,ink);rect(x+3,187,565,470,c);pic(img,x+14,199,543,244,ctr=ctr,frame=False);sticker(x+24,423,tag,milk,3,16);text(x+23,493,name,50);text(x+25,562,actor.upper(),17,ink,'m');body(x+25,608,[desc],21)
end()
# 03 / a film still with a book-cover style spine
begin(lime);meta('ПРОЛОГ  /  ПАРИЖ');pic('sowmestkadr1',41,118,697,442,-3,(.48,.48));sticker(87,543,'ВОТ ТАКОЕ СОБЕСЕДОВАНИЕ',pink,5,18);star(735,130,47,blue,.16);title(820,162,['Пришёл','за одной','подписью.'],45,59,418);rect(823,371,104,9,hot);body(821,421,['Дрисс получает работу','у Филиппа. Со временем','вместо «вы» и формальностей —','шутки, споры и дружба.'],22);checker(823,611,16,2,16,ink,milk);end()
# 04 & 05 / sleeve, turntable and a photograph of the actual scene
def music(bg,fg,track,ss,desc,img,time,accent,ctr):
    begin(bg);meta('СТОРОНА А   /   ДЕНЬ РОЖДЕНИЯ',fg);sticker(45,93,track,accent,3);title(43,211,ss,60,73,530,fg);body(48,435,desc,24,fg);record(1067,362,262,accent);pic(img,596,151,524,348,3,ctr);tc(time,633,555);star(1184,573,43,accent,.14);rect(0,665,1280,55,ink);text(44,685,'АУДИОДОРОЖКА  /  '+track,15,milk,'m');
    for i in range(81):
        h=6+21*abs(math.sin(i*.48)*math.cos(i*.14));rect(628+i*7.3,692-h/2,3,h,accent)
    end()
music(blue,milk,'ТРЕК 01  /  ДО',['Всё строго','по нотам.'],['На дне рождения звучит классика.','Гости слушают музыкантов.','Дриссу не хватает движения.'],'довключениямузыки','01:09:00–01:09:25',pink,(.50,.55))
music(hot,ink,'ТРЕК 02  /  ПОСЛЕ',['Дрисс меняет','пластинку.'],['Звучит Boogie Wonderland.','Дрисс не зовёт всех на танцпол —','он первым начинает танцевать.'],'послевключениямузыки','≈ 01:11:00–01:11:25',lime,(.38,.46))
# 06 / diptych
begin(milk);meta('СТОРОНА А   /   РАЗБОР');title(42,91,['Что изменил один трек?'],56,60,1185);rect(46,198,580,364,lilac);rect(654,198,580,364,lime);pic('довключениямузыки',64,216,537,220,ctr=(.5,.35),frame=False);pic('послевключениямузыки',671,216,537,220,ctr=(.4,.22),frame=False);text(65,473,'Зрители',43);text(673,473,'Участники',43);sticker(533,347,'ПЕРЕХОД →',pink,-8,19);body(48,604,['Личный пример помог гостям включиться. Но внезапная смена музыки могла','расстроить тех, кто пришёл ради классики.'],24);end()
# 07 / playlist on a giant concert ticket
begin(ink);record(1135,400,340,pink);meta('СТОРОНА А   /   ПАМЯТКА',milk);rect(49,120,989,489,lime);rect(72,142,944,445,lime);checker(49,120,54,2,18,ink,lime);title(84,189,['Плейлист для всех'],52,60,900);text(87,274,'БИЛЕТ НА ВЕЧЕР, ГДЕ ЕСТЬ ВЫБОР',17,ink,'m');
for i,(v,t) in enumerate([('Спросить','Что гости хотят услышать?'),('Чередовать','Концерт — а после него танцы.'),('Пригласить','Можно танцевать. Можно слушать.')]):
    y=342+i*74;line([(85,y-17),(986,y-17)],ink,1);text(88,y+4,f'0{i+1}',19,ink,'m');text(158,y,v,28);body(519,y+3,[t],22)
for y in [215,280,345,410,475,540]:ell(35,y,27,27,ink)
sticker(905,558,'ВХОД СВОБОДНЫЙ',pink,-9,22);end()
# 08 / gallery: museum wall label meets a meme
begin(lilac);meta('СТОРОНА Б   /   ГАЛЕРЕЯ');pic('первыйфрагментнепониманияискусства',46,119,717,427,2,(.5,.49));sticker(112,543,'ОДИН ХОЛСТ • ДВА МНЕНИЯ',lime,3);title(826,157,['Это точно','искусство?'],41,59,411);text(829,323,'ЭКСПОНАТ № 01',17,blue,'m');body(828,383,['Филипп всматривается','в абстрактную картину.','Дрисс смотрит на неё же —','и не понимает её цены.'],23);tc('≈ 00:31:20–00:31:45',828,583);end()
# 09 / paint marks
begin(mint);meta('СТОРОНА Б   /   МАСТЕРСКАЯ');
for j,c in enumerate([hot,lilac,lime,blue]):
    x=655+j*130;line([(x,582),(x+24,553),(x+111,183),(x+159,99)],c,46)
star(1176,121,57,hot,.18);sticker(46,116,'ЭКСПОНАТ № 02',pink,4);title(44,229,['Ладно.','Дайте кисть.'],57,72,558);body(48,447,['Вместо спора о чужой картине','Дрисс берётся за собственную.','Из зрителя — в автора.'],24);pic('созданиекартины',644,172,526,343,-4,(.5,.47));tc('≈ 00:59:57–01:00:20',701,590);end()
# 10 / auction receipt
begin(pink);meta('СТОРОНА Б   /   ПРОДАНО');rect(40,107,673,488,ink);pic('продажа картины',51,120,639,347,ctr=(.53,.38),frame=False);text(60,510,'Когда узнал цену:',33,milk);tc('≈ 01:16:00–01:16:25',69,618);rect(762,98,466,513,ink);rect(752,88,466,513,milk)
for i in range(18):poly([(752+i*26,600),(765+i*26,613),(778+i*26,600)],milk)
text(782,123,'ГАЛЕРЕЯ  /  ЧЕК № 001',17,ink,'m');line([(783,172),(1183,172)],ink);text(782,211,'11 000',78,mw=398);text(787,312,'ЕВРО',38);line([(784,378),(1183,378)],ink);body(786,415,['Филипп продал картину Дрисса.','Важнее цены то, что чью-то стену','теперь украсит его работа.'],20);rng=random.Random(11);x=785
while x<1178:
    w=rng.choice([2,3,5,7]);rect(x,529,w,40,ink);x+=w+rng.choice([2,4,6])
end()
# 11 / two gallery labels
begin(lime);meta('СТОРОНА Б   /   РАЗБОР');title(43,86,['Кисть вместо лекции'],59,65,1170);dots(0,578,1280,138,'#8EAD28')
for x,bg,lab,head,ss in [(50,pink,'01  /  ЧТО МЕШАЕТ','Навязывать',['Объяснять, что «надо» любить.','Так разговор о вкусах быстро','превращается в экзамен.']),(654,lilac,'02  /  ЧТО ПОМОГАЕТ','Предложить',['Спросить, что человек видит.','Дать ему кисть и холст —','а не список правильных ответов.'])]:
    rect(x+9,218,567,348,ink);rect(x,208,567,348,ink);rect(x+3,211,561,342,bg);sticker(x+20,234,lab,milk,2,15);text(x+24,327,head,40,mw=514);body(x+25,408,ss,23)
rect(44,614,1188,65,ink);text(68,638,'ИТОГ: ДРИСС НЕ ПРОСТО УЗНАЛ О ЖИВОПИСИ — ОН ПОПРОБОВАЛ САМ.',19,milk,'m');end()
# 12 / last cover, no call-to-action beyond reflection
begin(blue);meta('ОБРАТНАЯ СТОРОНА   /   РЕФЛЕКСИЯ',milk);sticker(44,96,'ТВОЯ ОЧЕРЕДЬ',pink,4);title(42,202,['Танцпол','или кисть?'],75,95,660,lime);pic('sowmestkadr3',770,104,407,348,-5,(.5,.47));star(1196,505,55,lime,.2);star(749,128,35,pink,.2);body(49,433,['Вам легче включиться, когда кто-то начинает','первым или когда можно попробовать самому?'],24,milk);rect(0,576,1280,144,pink);checker(0,576,80,1,16,ink,milk);text(48,625,'ВЫВОД',22,ink,'m');body(222,620,['Дрисс начал танцевать. Филипп поддержал художника. В обоих случаях','важен не «правильный» вкус, а возможность выбрать и поучаствовать.'],23);end()
R.core_properties.title='1+1. Дружба не по плейлисту';R.save(ROOT/'1plus1.pptx');pages[0].save(ROOT/'1plus1.pdf',save_all=True,append_images=pages[1:],quality=93,resolution=144)
if __name__=='__main__':
    print(f'Готово: {len(pages)} слайдов')
