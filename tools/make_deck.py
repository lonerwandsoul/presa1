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
    pages.append(im)
# Each spread is a self-contained visual joke. No timecodes or presenter prompts.
def full(name,center=(.5,.5)):
    im.paste(ImageOps.fit(Image.open(assets/(name+'.jpeg')).convert('RGB'),im.size,centering=center), (0,0))
def band(x,y,t,sz,bg,fg=ink,maxw=1170):
    while font(sz,'h').getlength(t)/K>maxw-40:sz-=1
    w=font(sz,'h').getlength(t)/K; rect(x,y,w+38,sz*1.36+12,bg);text(x+18,y+12,t,sz,fg)
def wave(x,y,n,c):
    for i in range(n):
        h=6+55*abs(math.sin(i*.41)*math.cos(i*.12));rect(x+9*i,y-h/2,5,h,c)
# 01: poster
begin(pink);ell(676,-93,778,778,lime);dots(927,0,365,716,ink);checker(0,652,80,4,18,ink,milk);text(35,102,'1+1',221);sticker(45,37,'ФРАНЦИЯ • 2011',milk,3,17);cutout(645,37,720);star(1170,114,59,hot,.18);star(709,129,44,blue,.22);title(43,337,['Когда вам','доверили','одну колонку'],56,73,623);rect(0,686,1280,34,ink);text(43,696,'ОЛИВЬЕ НАКАШ  ×  ЭРИК ТОЛЕДАНО',16,milk,'m');end()
# 02: opposing music profiles
begin(lilac);text(40,44,'Два админа плейлиста',63,mw=1195)
for x,img,name,actor,tag,desc,c,ctr in [(46,'philip','Филипп','Франсуа Клюзе','Ставит классику','Слышит первую скрипку.',pink,(.55,.36)),(654,'driss','Дрисс','Омар Си','Ставит фанк','Слышит, где вступает бас.',lime,(.6,.4))]:
    rect(x+9,154,571,511,ink);rect(x,145,571,511,ink);rect(x+3,148,565,505,c);pic(img,x+14,160,540,273,ctr=ctr,frame=False);sticker(x+29,406,tag,milk,4,18);text(x+22,487,name,49);text(x+24,557,actor.upper(),17,ink,'m');body(x+25,603,[desc],23)
end()
# 03: expectation / reality
begin(lime);pic('sowmestkadr1',-40,80,832,545,3,(.47,.48));star(760,110,64,pink,.15);sticker(64,602,'СОБЕСЕДОВАНИЕ',milk,4,20);title(864,129,['Зашёл','на пять','минут.'],47,60,374);rect(862,366,95,9,hot);title(863,424,['Вышел','в главной','роли.'],44,57,367);end()
# 04: a symphonic concert bill before the switch
begin(ink);record(1073,357,373,pink);pic('довключениямузыки',572,83,565,448,-3,(.50,.45));sticker(48,75,'ДЕНЬ РОЖДЕНИЯ ФИЛИППА',lilac,4,17);title(46,228,['Обещали','камерную','вечеринку.'],52,69,536,milk);band(47,536,'Привели оркестр.',31,lime,maxw=512);end()
# 05: screenshot-sized dance meme
begin(ink);full('послевключениямузыки',(.46,.4));band(685,52,'Я: сегодня без танцев.',38,milk,maxw=550);band(36,591,'Я через три ноты Boogie Wonderland:',37,lime,maxw=1192);end()
# 06: two cuts of the same dance become a reaction meme
begin(pink);text(42,51,'Стадии одного припева',58,mw=1180)
for x,img,t,c,ctr in [(43,'довключениямузыки','Я только послушаю.',milk,(.5,.42)),(652,'послевключениямузыки','Подержи мой пиджак.',lime,(.44,.36))]:
    rect(x+8,163,572,480,ink);rect(x,155,572,480,c);pic(img,x+10,165,552,360,ctr=ctr,frame=False);text(x+19,563,t,29,mw=535)
end()
# 07: faux vinyl single, the bonus track
begin(blue);record(987,370,364,lime);a=ImageOps.fit(Image.open(assets/'philip.jpeg'),sc(245,245),centering=(.64,.38));mask=Image.new('L',a.size);ImageDraw.Draw(mask).ellipse((0,0,a.width-1,a.height-1),255);im.paste(a,sc(864,247),mask);sticker(44,72,'БОНУС-ТРЕК',pink,4,19);title(41,190,['Филипп:','у меня дома','не клуб.'],57,78,614,milk);band(44,516,'Также Филипп: оставь.',29,lime,maxw=594);wave(53,647,58,pink);end()
# 08: museum scene, image crop removes embedded subtitles
begin(milk);src=Image.open(assets/'первыйфрагментнепониманияискусства.jpeg');src=src.crop((0,0,src.width,int(src.height*.78)));im.paste(ImageOps.fit(src,sc(1280,529)),(0,0));rect(0,529,1280,191,lilac);text(37,567,'Дрисс в галерее:',25,ink,'m');text(38,612,'«Я так тоже могу».',58,mw=1170);star(1138,623,57,hot,.15);end()
# 09: the punch line is the painting itself
begin(lime);sticker(45,44,'ДРИСС ДОМА',pink,4,20);title(42,180,['Опасная','стадия:'],60,81,509);band(42,398,'реально смог.',37,ink,milk,maxw=496)
for j,c in enumerate([hot,lilac,blue]):line([(627+170*j,640),(766+170*j,60)],c,44)
pic('созданиекартины',591,113,577,424,-4,(.52,.48));end()
# 10: sales receipt meme
begin(pink);pic('продажа картины',43,90,665,422,2,(.52,.37));band(44,576,'Моё лицо после первого заказа',25,ink,milk,maxw=680);rect(773,79,454,535,ink);rect(762,67,454,535,milk)
for i in range(18):poly([(762+i*25,601),(774+i*25,615),(787+i*25,601)],milk)
text(790,108,'ЧЕК 001  /  ЖИВОПИСЬ',17,ink,'m');line([(790,152),(1187,152)],ink);text(790,190,'«Я так тоже могу»',24,mw=392);text(792,258,'1 шт.',22,ink,'m');line([(790,314),(1187,314)],ink);text(790,345,'ИТОГО',18,ink,'m');text(784,390,'11 000 €',63,mw=404);rng=random.Random(11);x=792
while x<1184:
    w=rng.choice([2,3,5,7]);rect(x,520,w,44,ink);x+=w+rng.choice([2,4,6])
end()
# 11: old/new artist bio
begin(lilac);text(41,57,'Обновил описание профиля',53,mw=1190);pic('driss',40,158,587,430,-3,(.64,.5));rect(707,203,501,128,milk);text(731,225,'БЫЛО',17,ink,'m');text(730,266,'Не шарю в арте',30,mw=453);line([(735,286),(1158,281)],hot,5);rect(707,360,501,184,lime);text(732,386,'СТАЛО',17,ink,'m');title(731,428,['По заказам','в личку.'],38,50,446);star(1181,581,54,blue,.1);end()
# 12: reflection in a single visual callback
begin(ink);pic('sowmestkadr3',702,65,486,455,-5,(.49,.45));star(1178,598,54,lime,.19);sticker(47,70,'1+1',pink,5,22);title(45,191,['Плейлисты','так и не','совпали.'],56,76,636,milk);band(44,523,'Зато люди совпали.',40,lime,maxw=653);end()
R.core_properties.title='1+1. Когда вам доверили одну колонку';R.save(ROOT/'1plus1.pptx');pages[0].save(ROOT/'1plus1.pdf',save_all=True,append_images=pages[1:],quality=93,resolution=144)
