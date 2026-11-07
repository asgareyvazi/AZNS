#!/usr/bin/env python3
"""Seven reproducible source-scoped replacements for legacy portfolio figures.

All chart inputs and source anchors are retained in field_figure_observations.json.
They are exploratory evidence views, not verified schedules, costs or incidents.
"""
from pathlib import Path
import json,math,re,textwrap
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

D=Path('evidence');OUT=D/'field_figures';OUT.mkdir(exist_ok=True)
load=lambda n:json.loads((D/(n+'.json')).read_text())
program=load('program_summary_reviewed');tracks=load('drilling_track_observations');activities=load('activity_observations');fluid=load('fluid_observation_summary');bha=load('bha_observations');phases=load('program_phases_services_reviewed')
wmap={x['well']:x['source'] for x in load('well_bit_observations')}; order=['10','12','75','85','116','122WO1','122WO2','127WO1','127WO2','127ST2','166','207','40']
t={x['file']:x for x in tracks};a={x['file']:x for x in activities};f={x['well']:x for x in fluid};b={x['file']:x for x in bha};p={x['well']:x for x in program};ph={x['well']:x for x in phases}
BG='#f7f9fc';BLUE='#276987';ORANGE='#bf6d39';PURPLE='#7d5b97';INK='#233b50'

def canvas(title,subtitle,foot,w=12,h=8):
 fig=plt.figure(figsize=(w,h),dpi=300,facecolor=BG)
 fig.suptitle(title,x=.04,y=.975,ha='left',fontsize=16,fontweight='bold',color=INK)
 fig.text(.04,.925,subtitle,fontsize=9.1,color='#944b2f')
 fig.text(.04,.035,foot,fontsize=8.2,color='#944b2f')
 return fig

def save(fig,n):
 fig.canvas.draw();rend=fig.canvas.get_renderer()
 for text in fig.findobj(matplotlib.text.Text):
  if not text.get_visible() or not text.get_text() or text.get_clip_on():continue
  box=text.get_window_extent(rend)
  if box.width and box.height and (box.x0 < -2 or box.y0 < -2 or box.x1 > fig.bbox.x1+2 or box.y1 > fig.bbox.y1+2):raise ValueError((n,'clipped text',text.get_text(),box.bounds))
 fig.savefig(OUT/f'M{n}.png',dpi=300,facecolor=BG);plt.close(fig)

def main():
 metadata={'method':'Source-scoped chart inputs. Workbook observations are partial and may contain errors; programme plans are not actuals. Physical page numbers start at 1. No plotted zero substitutes for missing observations.','figures':{}}
 # Source-labelled single main/estimated duration, not a schedule or actual comparison.
 rows=[]
 for x in program:
  m=re.match(r'([0-9.]+)',x['duration_days_en']);assert m
  rows.append({'well':x['well'],'days':float(m.group(1)),'basis':x['duration_days_en'],'file':x['file'],'page':x['physical_page']})
 rows=sorted(rows,key=lambda x:x['days'],reverse=True);metadata['figures']['M1']=rows
 fig=canvas('Portfolio | Stated programme durations','Main-plan or estimated well durations as labelled; not actual days or a fleet timeline.','Sources: 14 reviewed programme information pages; page references appear at right. Scope and contingency require reconciliation.')
 ax=fig.add_axes([.15,.14,.68,.71]);ax.barh(range(14),[x['days'] for x in rows],color=BLUE,height=.65)
 ax.set_yticks(range(14),[x['well'] for x in rows]);ax.invert_yaxis();ax.set_xlim(0,175);ax.set_xlabel('Stated main or estimated well duration (days)');ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
 for i,x in enumerate(rows):ax.text(x['days']+2,i,f"{x['days']:g}  |  p.{x['page']}",va='center',fontsize=9)
 save(fig,1)
 # Inventory uses actual workbook sheet counts, not inferred daily records.
 inventory={x['file']:x for x in load('source_inventory')['files']}
 rows=sorted([{'well':w,'file':wmap[w],'sheets':inventory[wmap[w]]['sheets']} for w in order],key=lambda x:x['sheets'],reverse=True)
 metadata['figures']['M2']=rows
 fig=canvas('Portfolio | Original workbook sheet inventory','Sheet counts include all worksheet types; they are not report days or completeness.','Source: evidence/source_inventory.json. 127WO2 corresponds to AZNS127WOPack.xlsx; other labels use their pack identity.')
 ax=fig.add_axes([.15,.14,.72,.71]);ax.barh(range(len(rows)),[x['sheets'] for x in rows],color=BLUE,height=.65)
 ax.set_yticks(range(len(rows)),[x['well'] for x in rows]);ax.invert_yaxis();ax.set_xlim(0,390);ax.set_xlabel('Worksheets per original daily pack');ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0)
 for i,x in enumerate(rows):ax.text(x['sheets']+5,i,str(x['sheets']),va='center',fontsize=9)
 save(fig,2)
 # Raw status page coverage is independent of recognized data-page coverage.
 rows=[]
 for w in order:
  z=a[wmap[w]];rows.append({'well':w,'file':wmap[w],'status_pages':z['remark_days'],'complete_status_pages':sum(d['activity']['complete'] for d in z['days']),'recognized_data_pages':len(t[wmap[w]]['days']),'candidate_pairs':t[wmap[w]]['paired_sheet_candidates']})
 metadata['figures']['M4']=rows
 fig=canvas('Portfolio | Recognized daily-record coverage','Counts from distinct source-page searches; status pages and data pages are not equivalent.','Completeness means classified durations total 23.9–24.1 h; it does not verify timing, sequence or execution.')
 ax=fig.add_axes([.17,.14,.77,.70]);ys=range(len(rows))
 for delta,key,col,label in [(-.23,'status_pages','#9bb9c9','Status pages'),(0,'complete_status_pages',BLUE,'Complete status pages'),(.23,'recognized_data_pages',ORANGE,'Recognized data pages')]:
  ax.barh([i+delta for i in ys],[x[key] for x in rows],height=.21,color=col,label=label)
 ax.set_yticks(list(ys),[x['well'] for x in rows]);ax.invert_yaxis();ax.set_xlim(0,190);ax.set_xlabel('Recognized pages (counts, not actual duration)');ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0);ax.legend(ncol=3,loc='lower center',bbox_to_anchor=(.5,1.02),frameon=False,fontsize=8.5)
 save(fig,4)
 # Completeness per recognized data page; no status pages mixed into denominator.
 rows=[]
 for w in order:
  n=f[w]['recognized_data_sheets'];assert n>0
  vals={'BHA block':b[wmap[w]]['sheets_with_nonempty_bha_blocks'],'PV numeric':f[w]['fields']['pv']['numeric_accepted_count'],'Loss numeric':f[w]['fields']['downhole_loss']['numeric_accepted_count']}
  assert all(0<=v<=n for v in vals.values());rows.append({'well':w,'file':wmap[w],'data_pages':n,'counts':vals})
 metadata['figures']['M5']=rows
 fig=canvas('Portfolio | Selected data-field recognition','Share of recognized data pages with at least one BHA block, numeric PV, or numeric loss entry.','Missing or unrecognized is not zero loss or absence of equipment. Source readings, units and parser coverage remain unverified.')
 ax=fig.add_axes([.17,.14,.76,.69])
 for delta,key,col,mark in [(-.18,'BHA block',BLUE,'o'),(0,'PV numeric',PURPLE,'s'),(.18,'Loss numeric',ORANGE,'^')]:
  ax.scatter([100*x['counts'][key]/x['data_pages'] for x in rows],[i+delta for i in range(len(rows))],s=48,label=key,color=col,marker=mark)
 ax.set_yticks(range(len(rows)),[f"{x['well']}  (n={x['data_pages']})" for x in rows]);ax.invert_yaxis();ax.set_xlim(-3,105);ax.set_xticks(range(0,101,20));ax.set_xlabel('Recognized data pages with field (%)');ax.grid(axis='x',alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right','left']].set_visible(False);ax.tick_params(axis='y',length=0);ax.legend(ncol=3,loc='lower center',bbox_to_anchor=(.5,1.02),frameon=False)
 save(fig,5)
 # Daily report number is not an absolute timeline, MD not event or bit depth.
 rows=[]
 for well in ['207','40','127ST2','75']:
  points=[]
  for d in t[wmap[well]]['days']:
   z=d['fields']['md_end'];v=z and z['value']
   if v is not None and math.isfinite(v) and 0<=v<=6000:points.append({'day':d['day'],'md':v,'source':z})
  rows.append({'well':well,'file':wmap[well],'recognized_data_sheets':len(t[wmap[well]]['days']),'points':points})
 metadata['figures']['M7']=rows
 fig=canvas('Selected operations | Report-end depth observations','Four operations with varying recorded depth; dots are daily samples, not a field learning curve.','Worksheet prefix is not necessarily report number or elapsed time; report-end MD is not event depth.')
 axs=fig.subplots(2,2)
 for ax,row in zip(axs.flat,rows):
  pts=row['points'];ax.scatter([x['day'] for x in pts],[x['md'] for x in pts],s=11,c=BLUE,alpha=.7);ax.set_title(f"AZNS-{row['well']}  |  {len(pts)}/{row['recognized_data_sheets']} data pages",loc='left',fontsize=10,fontweight='bold',color=INK);ax.set_xlabel('Worksheet prefix');ax.set_ylabel('Report-end MD (m)');ax.invert_yaxis();ax.grid(alpha=.2);ax.spines[['top','right']].set_visible(False)
 fig.subplots_adjust(left=.10,right=.96,top=.84,bottom=.13,hspace=.50,wspace=.30);save(fig,7)
 # Programme-only selected phase abstracts, not chronology or actuals.
 ids=['122WO1','122WO2','127WO1','127WO2','127ST1','127ST2'];rows=[]
 short_phases={
  '122WO1':['Remove completion; mill / fish packer','Corrosion log; lap test; recomplete'],
  '122WO2':['Remove completion; mill / fish packer','Impression block and fishing assembly'],
  '127WO1':['Check and smooth 3,889–3,911 m','Contingency if interval is not passable'],
  '127WO2':['Wellhead checks; slickline operations','Test BOP; retrieve completion'],
  '127ST1':['Window 3,337 m; drill to 3,939 m','Liner-lap dry test; cement log'],
  '127ST2':['Whipstock 3,060 m; kickoff 3,055 m','Drill to 3,940 m; run / cement liner']}

 for w in ids:
  z=ph[w];rows.append({'well':w,'file':z['file'],'phases':z['selected_plan_phases'],'abridged_labels':short_phases[w]})
 metadata['figures']['M8']=rows
 fig=canvas('Wells 122 and 127 | Selected programme phases','Abridged programme statements, not a verified event sequence or proof of execution.','Each reference is a physical PDF page; review full programme scope and approved changes before linking operations.')
 ax=fig.add_axes([.04,.13,.92,.70]);ax.set_xlim(0,1);ax.set_ylim(len(rows),0);ax.axis('off')
 for i,row in enumerate(rows):
  ax.add_patch(plt.Rectangle((0,i),1,.96,facecolor='#e8f0f4' if i%2==0 else 'white',edgecolor='white'))
  ax.text(.015,i+.48,row['well'],fontsize=10,va='center',fontweight='bold',color=INK)
  for x,phase,short in zip([.14,.57],row['phases'],row['abridged_labels']):
   ax.text(x,i+.48,f"p.{phase['physical_page']}  {short}",fontsize=10,va='center',color=INK)
 save(fig,8)
 # Reconciliation issues are examples, not severity rankings.
 rows=[('40 casing','20-in depth 152.5 m vs 150 m','Programme pp.13 / 149'),('12 / 127WO1','Same 3,889–3,911 m collapse text','Programmes p.28 / p.58'),('127ST1 window','Sidetrack 3,300 m; window 3,337 m','Programme pp.11 / 70'),('122WO1 BHA','7-in scraper text vs 5-3/8-in OD','12-2 Data!I43 / K43'),('207 formation','Prognosed 3,915 m vs DDR 3,968 m','Programme p.16; 155-2 DDR Data!P86'),('85 rig-less','Zero rig time is not zero time or cost','CBS: AZNS-085 WO. Extra Work rows 41–49')]
 metadata['figures']['M9']=[{'scope':x,'issue':y,'source':z} for x,y,z in rows]
 fig=canvas('Portfolio | Examples requiring reconciliation','These are source differences or scope cautions; no risk severity or responsibility is assigned.','Depths, well identity, datum, timing, source context and approvals must be reconciled against original records.')
 ax=fig.add_axes([.04,.15,.92,.68]);ax.set_xlim(0,1);ax.set_ylim(len(rows),0);ax.axis('off')
 for i,(scope,issue,src) in enumerate(rows):
  ax.add_patch(plt.Rectangle((0,i),1,.96,facecolor='#e8f0f4' if i%2==0 else 'white',edgecolor='white'))
  ax.text(.015,i+.48,scope,fontsize=10,fontweight='bold',va='center',color=INK);ax.text(.20,i+.48,issue,fontsize=9,va='center',color=INK);ax.text(.65,i+.48,src,fontsize=8.5,va='center',color='#944b2f')
 save(fig,9)
 # Actual source fields on the same report, not simultaneous or formation attribution.
 rows=[]
 for well in order:
  pts=[]
  for d in t[wmap[well]]['days']:
   md=d['fields']['md_end'];loss=d['fields']['downhole_loss'];x=md and md['value'];y=loss and loss['value']
   if x is not None and y is not None and all(math.isfinite(v) for v in (x,y)) and 0<=x<=6000 and y>=0:pts.append({'day':d['day'],'md':md,'loss':loss})
  rows.append({'well':well,'file':wmap[well],'points':pts})
 metadata['figures']['M10']=rows
 fig=canvas('Portfolio | Source loss entries and report-end depth','Same-report numeric pairs only; depth is report-end MD, not a confirmed loss-event depth.','Individual values are not added; units, daily/cumulative basis and source quality require review. Missing wells are not zero loss.')
 ax=fig.add_axes([.11,.15,.68,.69]);colors=plt.get_cmap('tab10').colors;active=[r for r in rows if r['points']]
 for i,row in enumerate(active):
  ax.scatter([z['md']['value'] for z in row['points']],[z['loss']['value'] for z in row['points']],s=17,alpha=.72,color=colors[i%10],label=f"{row['well']} (n={len(row['points'])})")
 ax.set_yscale('symlog',linthresh=1);ax.set_xlabel('Report-end measured depth (m)');ax.set_ylabel('Source downhole / formation loss entry (bbl)');ax.grid(alpha=.2);ax.set_axisbelow(True);ax.spines[['top','right']].set_visible(False);ax.legend(loc='upper left',bbox_to_anchor=(1.03,1),frameon=False,title='Paired source entries',fontsize=9,title_fontsize=9)
 save(fig,10)
 (D/'field_figure_observations.json').write_text(json.dumps(metadata,indent=2,ensure_ascii=False)+'\n')
 print('Figures',sorted(metadata['figures']),'loss pair counts',[(x['well'],len(x['points'])) for x in rows])
if __name__=='__main__':main()
