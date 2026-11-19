"""Generate the 1+1 film-club slides (Pillow, python-pptx, DejaVu)."""
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageFilter, ImageChops
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pathlib import Path
from io import BytesIO
from functools import lru_cache
import math, random, zipfile, xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[1]; assets=ROOT/'pres'; K=1.5; W,H=1280,720
ink='#142C35';milk='#FAF5EA';pink='#FF86C8';hot='#FF408F';lime='#D6FF48';blue='#4138D3';lilac='#C2B3FF';orange='#FF8353'; mint='#A0EECC'
R=Presentation();R.slide_width=Inches(13+1/3);R.slide_height=Inches(7.5);pages=[]
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
def end():
    # Keep the designed artwork identical on machines without the display font.
    s=R.slides.add_slide(R.slide_layouts[6]);buf=BytesIO();im.save(buf,'PNG',optimize=True);buf.seek(0);s.shapes.add_picture(buf,0,0,R.slide_width,R.slide_height)
    for x,y,t,sz,c in captions:
        w=font(sz,'b').getlength(t)/K; assert x+w<1278,(t,w)
        a=s.shapes.add_textbox(Inches(x/96),Inches((y-sz*.18)/96), Inches((w+34)/96), Inches(sz*1.6/96));fr=a.text_frame;fr.word_wrap=False;fr.margin_top=fr.margin_left=fr.margin_right=fr.margin_bottom=0;p=fr.paragraphs[0];run=p.add_run();run.text=t;run.font.size=Pt(sz*.75);run.font.name='DejaVu Sans';run.font.color.rgb=RGBColor.from_string(c[1:]);text(x,y,t,sz,c,'b')
    pages.append(im);s.notes_slide.notes_text_frame.text=notes[len(pages)]
# Notes follow the supplied speaking script. The slide surface is deliberately concise.
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};notes={};n=0
with zipfile.ZipFile(next(assets.glob('*.docx'))) as z:root=ET.fromstring(z.read('word/document.xml'))
for p in root.findall('.//w:p',ns):
    t=''.join(e.text or '' for e in p.findall('.//w:t',ns))
    if t.startswith('Слайд '):n=int(t.split()[1].rstrip('.'));notes[n]=''
    elif t=='ПОРЯДОК ДЕЙСТВИЙ':break
    elif n and t.startswith(('Сказать:', 'Видео №', 'Рефлексия:', 'Вопрос аудитории:', 'Заключительный вопрос:', 'Выслушать')):notes[n]+=('\n\n' if notes[n] else '')+t.replace('Видео №','Ориентир в фильме №')
aqua='#A3DDCA';orange='#FF9869';stone='#DADACF';ochre='#E6B84E';teal='#215C59'
def mast(section,fg=ink):
    text(43,32,'1+1',22,fg);text(135,35,section.upper(),16,fg,'m');text(1131,35,f'{len(pages)+1:02} / 12',16,fg,'m');line([(43,72),(1236,72)],fg,1)
def image(name,x,y,w,h,ctr=(.5,.5),bottom=1):
    a=Image.open(assets/(name+'.jpeg')).convert('RGB');a=a.crop((0,0,a.width,round(a.height*bottom)));im.paste(ImageOps.fit(a,sc(w,h),method=Image.Resampling.LANCZOS,centering=ctr),sc(x,y))
def label(x,y,t,bg):
    w=font(16,'m').getlength(t)/K+28;rect(x,y,w,33,bg);text(x+14,y+9,t,16,ink,'m')
def plus(x,y,sz,c):rect(x-sz/2,y-sz*.105,sz,sz*.21,c);rect(x-sz*.105,y-sz/2,sz*.21,sz,c)
def wax(x,y,r,c):
    ell(x-r,y-r,2*r,2*r,ink)
    for v in (.46,.58,.68,.77,.85,.92):ell(x-r*v,y-r*v,r*v*2,r*v*2,None,'#536870',1)
    ell(x-r*.3,y-r*.3,r*.6,r*.6,c);ell(x-4,y-4,8,8,ink)
# I. People
begin(ink);ell(699,-40,730,730,orange);cutout(649,34,716);text(43,44,'ФРАНЦИЯ   /   2011',17,milk,'m');text(31,114,'1+1',223,milk);title(47,391,['Дружба на слух','и в цвете'],52,67,647,orange);body(49,569,['Как Филипп и Дрисс знакомят друг друга','с музыкой и живописью'],24,milk);rect(0,669,1280,51,ink);text(48,689,'РЕЖИССЁРЫ   /   ОЛИВЬЕ НАКАШ И ЭРИК ТОЛЕДАНО',16,milk,'m');end()
begin(milk);mast('Герои');text(41,106,'Два взгляда на культуру',52,mw=1190)
for x,name,hero,actor,ss,c in [(43,'philip','Филипп','Франсуа Клюзе',['Ценит камерную музыку и живопись.','Нанимает Дрисса помощником.'],orange),(655,'driss','Дрисс','Омар Си',['Говорит прямо, любит фанк.','На празднике первым выйдет танцевать.'],aqua)]:
    image(name,x,198,579,260,(.58,.39));rect(x,458,579,87,c);text(x+17,478,hero,37);text(x+19,568,actor.upper(),17,teal,'m');body(x+18,611,ss,22)
end()
begin(ink);mast('Знакомство',milk);image('продажа картины',43,108,705,485,(.47,.45));label(43,593,'— И ДОЛГО ЭТО БУДЕТ?   — ЧЕТЫРЕ ЧАСА.',orange);title(790,143,['Началось','с работы.'],49,65,436,milk);body(792,337,['Филипп ищет помощника,','Дриссу нужна подпись.','Они начинают проводить время','вместе — и становятся друзьями.'],23,milk);rect(792,511,64,7,orange);body(792,555,['Оперу они слушают по-разному.','Смеются в ложе уже вместе.'],21,orange);end()
# II. Involvement in a shared event
begin(milk);mast('01   /   День рождения');image('довключениямузыки',539,98,701,568,(.56,.43));label(48,141,'ДО   /   КЛАССИКА',orange);title(43,233,['Вечер','по нотам'],57,76,470);body(48,453,['В честь Филиппа выступают','музыканты. Гости слушают,','но Дриссу не хватает движения.'],23);end()
begin(ink);mast('01   /   День рождения',milk);wax(249,413,239,orange);image('послевключениямузыки',43,146,557,366,(.40,.42));label(43,512,'BOOGIE WONDERLAND',orange);title(653,156,['Дрисс выходит','танцевать'],49,65,583,milk);body(657,353,['Он меняет музыку и начинает сам.','Гостям не нужно ждать приглашения:','они видят движение и подхватывают его.'],23,milk);rect(656,535,524,2,orange);text(656,567,'ПОСЛЕ   /   ТАНЦПОЛ',17,orange,'m');end()
begin(milk);mast('01   /   Что сработало');title(43,108,['Приглашение без слов'],54,69,1192)
for x,img,c,head,ss in [(43,'довключениямузыки',stone,'Риск',['Музыку сменили без спроса. Кому-то','могло хотелось дослушать классику.']),(655,'послевключениямузыки',orange,'Результат',['Гости включились в танец,','а Филипп улыбается.'])]:
    if head=='Риск':ss=['Музыку сменили без спроса. Кто-то','мог хотеть дослушать классику.']
    image(img,x,213,579,220,(.48,.31));rect(x,433,579,179,c);text(x+18,454,head,33);body(x+19,517,ss,22)
body(46,655,['Дрисс предложил гостям не просто слушать музыку, а двигаться вместе с ним.'],22);end()
begin(orange);mast('01   /   Вывод для организатора');title(42,152,['На одном','празднике —','разные гости'],51,72,606);wax(121,673,108,milk)
for i,(h,t) in enumerate([('До вечера','Узнать, что любят слушать гости.'),('В программе','После концерта предложить танцы.'),('На танцполе','Пригласить всех, не тянуть за руку.')]):
    y=144+i*160;line([(688,y),(1233,y)],ink,2);text(690,y+24,'0'+str(i+1),17,ink,'m');text(756,y+19,h,31,mw=463);body(756,y+80,[t],22)
end()
# III. From looking at art to making it
begin(milk);mast('02   /   В галерее');image('первыйфрагментнепониманияискусства',43,105,710,491,(.49,.47),.79);label(43,596,'АБСТРАКТНАЯ ЖИВОПИСЬ',aqua);title(791,175,['Один холст,','два взгляда'],40,60,443);body(793,367,['Филипп видит в картине ценность.','Дрисса удивляет цена: ему кажется,','что так нарисовать легко.'],22);plus(1193,610,69,aqua);end()
begin(aqua);mast('02   /   В мастерской');title(43,155,['Из зрителя','в автора'],56,79,536);body(48,388,['После похода в галерею Дрисс','сам берётся за кисть. Теперь он','не оценивает чужую работу,','а решает, что делать со своей.'],23);rect(623,139,568,406,teal);image('созданиекартины',642,117,578,407,(.49,.43));label(644,547,'РАБОТА ДРИССА',milk);end()
begin(ink);mast('02   /   Признание',milk);image('продажа картины',43,124,609,425,(.48,.46));text(44,572,'ФИЛИПП СООБЩАЕТ О ПРОДАЖЕ',16,milk,'m');text(700,193,'11 000',96,aqua,mw=524);text(708,317,'ЕВРО',40,milk);line([(710,408),(1220,408)],aqua);body(710,456,['За столько Филипп продал картину Дрисса.','Ещё недавно Дрисс шутил о живописи,','а теперь у его работы есть покупатель.'],22,milk);end()
begin(milk);mast('02   /   Вывод');title(43,115,['Как поддержать интерес'],53,65,1195);rect(43,221,527,352,stone);rect(591,221,644,352,aqua);text(66,252,'РИСК',17,ink,'m');title(64,320,['Объяснять,','что любить'],38,54,470);body(67,460,['Спор о «правильном» вкусе легко','отбивает желание продолжать.'],22);text(618,252,'ДРУГОЙ ПУТЬ',17,ink,'m');title(616,320,['Дать кисть','и свободу'],38,54,584);body(619,460,['Начать с того, что человек уже заметил.','Предложить попробовать, не ставя оценки.'],22);body(44,631,['Дрисса заинтересовал не разговор о технике, а возможность самому смешивать краски.'],23);end()
begin(teal);mast('Рефлексия',milk);image('sowmestkadr3',785,118,448,410,(.5,.4));title(42,158,['Не обязательно','любить одно','и то же.'],57,73,708,milk);body(46,446,['Дрисс первым вышел танцевать.','Филипп всерьёз отнёсся к картине друга.'],25,milk);line([(44,577),(1233,577)],aqua,2);body(47,617,['В обоих эпизодах один открыл другому возможность, но оставил за ним выбор.'],25,milk);end()
R.core_properties.title='1+1. Дружба на слух и в цвете';R.save(ROOT/'1plus1.pptx');pages[0].save(ROOT/'1plus1.pdf',save_all=True,append_images=pages[1:],quality=93,resolution=144)
