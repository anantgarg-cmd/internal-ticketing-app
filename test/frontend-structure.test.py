from html.parser import HTMLParser
from pathlib import Path

VOID={'base','br','col','embed','hr','img','input','link','meta','param','source','track','wbr'}
class Validator(HTMLParser):
 def __init__(self): super().__init__(convert_charrefs=True); self.stack=[]; self.errors=[]; self.ids=[]; self.nodes=[]; self.seq=0
 def handle_starttag(self,tag,attrs):
  attrs=dict(attrs); parent=self.stack[-1] if self.stack else None; node=(tag,attrs,parent,self.seq); self.seq+=1; self.nodes.append(node)
  if 'id' in attrs:self.ids.append(attrs['id'])
  if tag not in VOID:self.stack.append(node)
 def handle_startendtag(self,tag,attrs): self.handle_starttag(tag,attrs); self.stack=self.stack[:-1] if self.stack and self.stack[-1][0]==tag else self.stack
 def handle_endtag(self,tag):
  if not self.stack or self.stack[-1][0]!=tag:self.errors.append(f'unmatched closing </{tag}>; open={[x[0] for x in self.stack[-4:]]}')
  else:self.stack.pop()
v=Validator();v.feed(Path('Index.html').read_text());v.close()
assert not v.errors, v.errors
assert not v.stack, 'unclosed tags: '+str([x[0] for x in v.stack])
assert len(v.ids)==len(set(v.ids)), 'duplicate ids: '+str({x for x in v.ids if v.ids.count(x)>1})
def one(tag=None,id=None,cls=None):
 found=[n for n in v.nodes if (tag is None or n[0]==tag) and (id is None or n[1].get('id')==id) and (cls is None or cls in n[1].get('class','').split())];assert len(found)==1,(tag,id,cls,len(found));return found[0]
layout=one('div',cls='app-layout'); nav=one('nav','nav'); main=one('main')
assert nav[2] is layout and main[2] is layout
pages=['registrationPage','raisePage','myPage','queuePage','numbersPage']
for page in pages: assert one('section',page)[2] is main
raise_node=one('section','raisePage')
for page in pages[2:]: assert one('section',page)[2] is not raise_node
for host in ['activityTicketList','activityTicketPagination','queueTickets','queuePagination','numbers','detailContent','duplicateList','reopenContent']:one(id=host)
for modal in ['duplicateModal','detailModal','reopenModal']:
 n=one('div',modal);assert n[2] is not main
scripts=[n for n in v.nodes if n[0]=='script'];assert scripts and scripts[-1][3]>max(n[3] for n in v.nodes if n[1].get('id') in ['globalProcessingLoader','toastStack'])
print(f'frontend structure: {len(v.nodes)} elements, {len(v.ids)} unique IDs, {len(pages)} direct pages')
