#!/usr/bin/env python3
"""Check paragraph line counts and preferred line endings on the correct physical PDF page.

Provide a proofread manifest (--manifest), or line-preserving text with #page N and #block ID
markers (--scan-lines), or OCR paragraphs (--scan-tsv). OCR and unmarked text are provisional:
--verified-reference means the agent has checked paragraph boundaries and transcription against
all relevant scan images. A success checks declared blocks only, not unlisted paper content.
Vertical spacing (manifest only, all optional, in points measured from the scan at the rendered
page/column scale): a block's `spacing.observed_baseline_pitch_pt` is compared with the rendered
baseline pitch inside the paragraph, and `spacing.observed_baseline_to_next_block_pt` with the
rendered distance from its last baseline to the first baseline of the next listed block on the same
page. A page's `answer_line_count` and optional `answer_line_pitch_pt` are compared with the rendered
rows of typed underscores. A requested measurement that cannot be made is incomplete, never a pass.
Requires pdfplumber only for reading the rendered PDF.
"""
import argparse
import csv
import difflib
import json
import math
from pathlib import Path
import re
import sys
import unicodedata


def is_cjk(char):
    n = ord(char)
    return 0x3400 <= n <= 0x9FFF or 0x20000 <= n <= 0x323AF or 0x3040 <= n <= 0x30FF or 0xAC00 <= n <= 0xD7AF


def tokens(text):
    text = unicodedata.normalize('NFKC', text).casefold().translate(str.maketrans({'‘': "'", '’': "'", '“': '"', '”': '"', '–': '-', '−': '-'}))
    result, word = [], ''
    for char in text:
        category = unicodedata.category(char)
        if not is_cjk(char) and category[0] in 'LNM':
            word += char
            continue
        if word:
            result.append(word); word = ''
        if is_cjk(char) or category[0] == 'S' or char in '+-=<>/%':
            result.append(char)
    if word:
        result.append(word)
    return result


def load_text(path):
    blocks, page, block_id, lines, explicit = [], 1, None, [], False
    def flush():
        if lines:
            blocks.append({'id': block_id or f'p{page}-auto{len(blocks)+1}', 'page': page,
                           'lines': list(lines), 'explicit_boundaries': block_id is not None})
            lines.clear()
    for raw in Path(path).read_text(encoding='utf-8-sig').splitlines():
        if re.match(r'^#page(?:\s|$)', raw):
            flush();block_id=None
            parts=raw.split()
            page=int(parts[1]) if len(parts)>1 else page+1
        elif re.match(r'^#block(?:\s|$)', raw):
            flush();block_id=raw[6:].strip()
            if not block_id:raise ValueError('#block requires a stable paragraph ID')
            explicit=True
        elif not raw.strip():
            flush();block_id=None
        elif not raw.startswith('#'):
            lines.append(raw)
    flush()
    return blocks, explicit and all(b['explicit_boundaries'] for b in blocks), None, False


def load_tsv(paths, min_conf):
    blocks=[]
    for page,path in enumerate(paths,1):
        groups={}
        with Path(path).open(encoding='utf-8-sig') as stream:
            for row in csv.DictReader(stream,delimiter='\t',quoting=csv.QUOTE_NONE):
                if row.get('level')!='5' or not row.get('text','').strip():continue
                if float(row['conf'])<min_conf:continue
                key=(int(row['block_num']),int(row['par_num']))
                groups.setdefault(key,{}).setdefault(int(row['line_num']),[]).append((int(row['left']),row['text']))
        for key,lines in sorted(groups.items()):
            blocks.append({'id':f'p{page}-ocr{key[0]}-{key[1]}','page':page,
                'lines':[' '.join(t for _,t in sorted(words)) for _,words in sorted(lines.items())]})
    return blocks, True, len(paths), False


def load_manifest(path):
    source=json.loads(Path(path).read_text(encoding='utf-8-sig'));blocks=[]
    if not isinstance(source,dict):raise ValueError('Manifest root must be an object')
    if not isinstance(source.get('reference_verified',False),bool):raise ValueError('reference_verified must be a JSON boolean')
    if 'pages' in source:
        pages=source['pages']
        if not isinstance(pages,list):raise ValueError('pages must be a list')
        for page in pages:
            if not isinstance(page,dict) or not isinstance(page.get('blocks',[]),list):raise ValueError('Each page must be an object with a blocks list')
            for block in page.get('blocks',[]):
                if not isinstance(block,dict):raise ValueError('Each paragraph block must be an object')
                blocks.append({**block,'page':page.get('physical_page',page.get('page'))})
        count=len(pages)
        numbers=[p.get('physical_page',p.get('page')) for p in pages]
        if any(type(n) is not int for n in numbers) or numbers!=list(range(1,count+1)):raise ValueError('Manifest pages must use consecutive physical indices starting at 1; printed labels are separate')
    else:
        blocks=source.get('blocks',[]);count=source.get('physical_page_count')
    if count is not None and (type(count) is not int or count<1):raise ValueError('physical_page_count must be a positive integer')
    return blocks, True, count, source.get('reference_verified',False)


def pdf_lines(page,bbox=None):
    return [text for text,_ in pdf_rows(page,bbox)]


def pdf_rows(page,bbox=None):
    """Rendered (text, baseline) rows; the baseline is the bottom of the line's full-height glyphs. Characters join a line when they overlap it vertically by at least half
    their height, so real subscripts/superscripts (S1, x2, 10^3) stay on their printed line instead
    of forming a separate line, and a letter plus its subscript join into one word."""
    if bbox is not None:
        if len(bbox)!=4:raise ValueError('render_bbox must be [left, top, right, bottom] in rendered PDF points')
        page=page.crop(tuple(bbox))
    rows=[]
    for char in sorted(page.chars,key=lambda c:c['top']-c['bottom']):  # tallest first define the lines
        top,bottom=char['top'],char['bottom'];height=bottom-top
        if height<=0 or char.get('text') is None:continue
        best,overlap=None,0
        for row in rows:
            shared=min(bottom,row['bottom'])-max(top,row['top'])
            if shared>=0.5*min(height,row['bottom']-row['top']) and shared>overlap:best,overlap=row,shared
        if best is None:rows.append({'top':top,'bottom':bottom,'chars':[char]})
        else:best['chars'].append(char)
    lines=[]
    for row in sorted(rows,key=lambda r:r['top']):
        text,previous='',None
        for char in sorted(row['chars'],key=lambda c:c['x0']):
            if char['text'].isspace():text+=' '
            else:
                if previous is not None and char['x0']-previous['x1']>0.2*min(char.get('size',10),previous.get('size',10)):text+=' '
                text+=char['text']
            previous=char
        text=' '.join(text.split())
        if text:lines.append((text,max(c['bottom'] for c in row['chars'] if c['bottom']-c['top']>=0.9*(row['bottom']-row['top']))))
    return lines


def hyphenation_variants(scan_lines, confirmed_breaks=None):
    """Relax only independently confirmed discretionary hyphens, using 1-based line indices.

    Lower-case continuation alone cannot distinguish consumption from cost-effective or re-sign.
    Printed/lexical hyphens stay strict unless the source manifest explicitly confirms the break.
    """
    printed=[tokens(line) for line in scan_lines]
    confirmed=[] if confirmed_breaks is None else confirmed_breaks
    if not isinstance(confirmed,list) or any(type(n) is not int for n in confirmed) or len(set(confirmed))!=len(confirmed):
        raise ValueError('discretionary_hyphen_breaks must be a list of unique 1-based line integers')
    breaks=[]
    for number in confirmed:
        i=number-1
        if not 0<=i<len(printed)-1 or len(printed[i])<2 or printed[i][-1]!='-' or not printed[i][-2].isalpha() or not printed[i+1] or not printed[i+1][0].isalpha():
            raise ValueError('Each discretionary_hyphen_breaks entry must follow a printed word- line with an alphabetic continuation')
        breaks.append(i)
    yield 'as_printed',printed
    if not breaks:return
    split=[list(t) for t in printed]
    for i in breaks:split[i]=split[i][:-1]
    yield 'hyphen_dropped',split
    joined=[list(t) for t in split]
    for i in sorted(breaks,reverse=True):
        if joined[i+1]:joined[i][-1]+=joined[i+1][0];joined[i+1]=joined[i+1][1:]
    yield 'joined_word',joined


def compare_block(block, rendered_lines, diagnostics=True, baselines=None):
    scan_lines=[str(s) for s in block.get('lines',[]) if str(s).strip()]
    render_tokens=[]
    for row,line in enumerate(rendered_lines):
        render_tokens.extend((t,row) for t in tokens(line))
    flat=[t for t,_ in render_tokens]
    source={'id':block['id'],'page':block['page'],'scan_line_count':len(scan_lines)}
    for variant,line_tokens in hyphenation_variants(scan_lines,block.get('discretionary_hyphen_breaks')):
        query=[t for line in line_tokens for t in line];length=len(query)
        starts=[i for i in range(len(flat)-length+1) if flat[i:i+length]==query] if query else []
        if starts:break
    if variant!='as_printed' and not starts:
        variant,line_tokens=next(hyphenation_variants(scan_lines,block.get('discretionary_hyphen_breaks')))
        query=[t for line in line_tokens for t in line];length=len(query)
    if not query:return {**source,'status':'empty_reference','line_count_matches':None}
    if not starts:
        result={**source,'status':'unmatched_text','line_count_matches':None,'scan_text':' / '.join(scan_lines)[:200]}
        # Diagnostics only (the block stays unresolved): show the closest rendered text and the
        # token differences, so a transcription slip or a real wording error can be fixed.
        longest=difflib.SequenceMatcher(None,flat,query,autojunk=False).find_longest_match(0,len(flat),0,length) if diagnostics else None
        if longest is not None and longest.size:
            start=max(0,longest.a-longest.b-5);window=flat[start:start+length+10]
            matcher=difflib.SequenceMatcher(None,query,window,autojunk=False)
            codes=[c for c in matcher.get_opcodes() if c[0]!='equal']
            # Rendered text beyond either end of the reference is context, not a difference.
            codes=[c for c in codes if not (c[0]=='insert' and c[1] in (0,length))]
            result['closest_render_similarity']=round(max(0.0,min(1.0,1-sum(max(i2-i1,j2-j1) for _,i1,i2,j1,j2 in codes)/length)),3)
            result['differences_first_5']=[{'reference':' '.join(query[i1:i2]),'render':' '.join(window[j1:j2])}
                                           for _,i1,i2,j1,j2 in codes][:5]
        return result
    if len(starts)>1:return {**source,'status':'ambiguous_text','line_count_matches':None,
                            'matches_on_page':len(starts),'hint':'Provide a render_bbox for this paragraph'}
    start=starts[0];rows=[row for _,row in render_tokens[start:start+length]]
    actual=len(set(rows));scan_breaks=set();offset=0
    for line in line_tokens[:-1]:
        offset+=len(line);scan_breaks.add(offset)
    render_breaks={i for i in range(1,length) if rows[i]!=rows[i-1]}
    result={**source,'status':'matched','render_line_count':actual,'line_count_matches':len(scan_lines)==actual,
            # A joined word means the scan broke inside a word the render keeps whole: endings differ.
            'exact_line_endings':scan_breaks==render_breaks and variant!='joined_word',
            'render_text':' / '.join(rendered_lines[min(rows):max(rows)+1])[:240]}
    if variant!='as_printed':
        result['hyphenation_variant']=variant
        result['confirmed_discretionary_hyphen_breaks']=block['discretionary_hyphen_breaks']
    if baselines is not None:
        ys=[baselines[r] for r in sorted(set(rows))]
        result['first_baseline_pt']=round(ys[0],2);result['last_baseline_pt']=round(ys[-1],2)
        steps=[b-a for a,b in zip(ys,ys[1:])]
        result['render_baseline_steps_pt']=steps
        result['render_baseline_pitch_pt']=round(median(steps),2) if steps else None
    return result


def median(values):
    values=sorted(values);n=len(values)
    return (values[n//2] if n%2 else (values[n//2-1]+values[n//2])/2) if n else None


def check_close(expected,actual,tolerance):
    if actual is None:return {'expected':expected,'actual':None,'ok':None}
    return {'expected':expected,'actual':round(actual,2),'ok':abs(actual-expected)<=tolerance}


def check_pitch(expected,steps,tolerance):
    # A median alone can hide one large gap or alternating wrong intervals.
    result=check_close(expected,median(steps),tolerance)
    result['steps_pt']=[round(step,2) for step in steps]
    if steps:result['ok']=all(math.isfinite(step) and abs(step-expected)<=tolerance for step in steps)
    return result


def number(value,name,allow_none=True):
    if value is None and allow_none:return None
    if type(value) not in (int,float) or not math.isfinite(value) or value<0:raise ValueError(name+' must be a non-negative number of points')
    return value


def load_page_spacing(path):
    """Physical page -> {'count', 'pitch'} for typed answer lines declared in the manifest."""
    source=json.loads(Path(path).read_text(encoding='utf-8-sig'));items=[]
    if isinstance(source.get('pages'),list):
        items=[p for p in source['pages'] if isinstance(p,dict)]
    if isinstance(source.get('answer_lines'),list):
        items+= [p for p in source['answer_lines'] if isinstance(p,dict)]
    specs={}
    for item in items:
        if item.get('answer_line_count') is None:
            if item.get('answer_line_pitch_pt') is not None:raise ValueError('answer_line_pitch_pt requires answer_line_count')
            continue
        page=item.get('physical_page',item.get('page'))
        if type(page) is not int or page<1 or type(item['answer_line_count']) is not int or item['answer_line_count']<0:
            raise ValueError('answer_line_count needs a non-negative integer and an integer physical_page')
        if page in specs:raise ValueError('Answer-line specifications must be unique per physical page')
        specs[page]={'count':item['answer_line_count'],'pitch':number(item.get('answer_line_pitch_pt'),'answer_line_pitch_pt')}
    return specs


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('render_pdf',type=Path)
    source=parser.add_mutually_exclusive_group(required=True)
    source.add_argument('--manifest',type=Path)
    source.add_argument('--scan-lines',type=Path)
    source.add_argument('--scan-tsv',nargs='+',type=Path)
    parser.add_argument('--verified-reference',action='store_true',help='Agent has proofread the reference and paragraph boundaries against the scan')
    parser.add_argument('--min-conf',type=float,default=60)
    parser.add_argument('--pitch-tolerance',type=float,default=1.0,help='Allowed baseline-pitch difference in points (default 1.0)')
    parser.add_argument('--gap-tolerance',type=float,default=3.0,help='Allowed block-to-block baseline difference in points (default 3.0)')
    parser.add_argument('--json',type=Path,help='New JSON report path; existing files are never overwritten')
    args=parser.parse_args()
    try:
        number(args.pitch_tolerance,'--pitch-tolerance',allow_none=False)
        number(args.gap_tolerance,'--gap-tolerance',allow_none=False)
        if args.manifest:blocks,explicit,page_count,verified=load_manifest(args.manifest)
        elif args.scan_lines:blocks,explicit,page_count,verified=load_text(args.scan_lines)
        else:blocks,explicit,page_count,verified=load_tsv(args.scan_tsv,args.min_conf)
        verified=verified or args.verified_reference
        if not isinstance(blocks,list):raise ValueError('blocks must be a list')
        for block in blocks:
            if not isinstance(block,dict):raise ValueError('Each paragraph block must be an object')
            if type(block.get('page')) is not int or block['page']<1:raise ValueError('Block page must be a positive physical-page integer')
            if not isinstance(block.get('id'),str) or not block['id'].strip():raise ValueError('Every block needs a nonempty string ID')
            if not isinstance(block.get('lines'),list) or not all(isinstance(s,str) for s in block['lines']):raise ValueError('lines must be a list of printed-line strings')
            list(hyphenation_variants([line for line in block['lines'] if line.strip()],block.get('discretionary_hyphen_breaks')))
            spacing=block.get('spacing')
            if spacing is not None:
                if not isinstance(spacing,dict):raise ValueError('spacing must be an object')
                number(spacing.get('observed_baseline_pitch_pt'),'observed_baseline_pitch_pt')
                number(spacing.get('observed_baseline_to_next_block_pt'),'observed_baseline_to_next_block_pt')
            box=block.get('render_bbox')
            if box is not None and (not isinstance(box,list) or len(box)!=4 or any(type(v) not in (int,float) or not math.isfinite(v) for v in box) or box[0]>=box[2] or box[1]>=box[3]):raise ValueError('render_bbox requires four finite coordinates with positive width and height')
        ids=[b['id'] for b in blocks]
        if len(set(ids))!=len(ids):raise ValueError('Every block needs a unique paragraph ID')
        page_spacing=load_page_spacing(args.manifest) if args.manifest else {}
        import pdfplumber
        with pdfplumber.open(args.render_pdf) as pdf:
            results,page_cache=[],{}
            def page_rows(number):
                if number not in page_cache:page_cache[number]=pdf_rows(pdf.pages[number-1])
                return page_cache[number]
            def page_lines(number):
                return [text for text,_ in page_rows(number)]
            for block in blocks:
                page=block['page']
                if not 1<=page<=len(pdf.pages):
                    results.append({'id':block['id'],'page':page,'status':'missing_page','line_count_matches':None});continue
                rows=pdf_rows(pdf.pages[page-1],block['render_bbox']) if block.get('render_bbox') else page_rows(page)
                result=compare_block(block,[t for t,_ in rows],baselines=[y for _,y in rows])
                if result['status']=='unmatched_text':
                    # Still unresolved: text on the wrong rendered page means a page-allocation error.
                    elsewhere=[n for n in range(1,len(pdf.pages)+1) if n!=page
                               and compare_block(block,page_lines(n),diagnostics=False)['status'] in ('matched','ambiguous_text')]
                    if elsewhere:result['found_on_other_render_pages']=elsewhere
                results.append(result)
            actual_pages=len(pdf.pages)
            # Vertical spacing: only the measurements the manifest declares.
            vertical=[]
            for index,(block,result) in enumerate(zip(blocks,results)):
                spacing=block.get('spacing') or {};checks={}
                if spacing.get('observed_baseline_pitch_pt') is not None:
                    checks['baseline_pitch']=check_pitch(spacing['observed_baseline_pitch_pt'],result.get('render_baseline_steps_pt',[]),args.pitch_tolerance)
                if spacing.get('observed_baseline_to_next_block_pt') is not None:
                    following=next(((b,r) for b,r in zip(blocks[index+1:],results[index+1:]) if b['page']==block['page']),None)
                    distance=(following[1]['first_baseline_pt']-result['last_baseline_pt']
                              if following and following[1].get('first_baseline_pt') is not None and result.get('last_baseline_pt') is not None else None)
                    checks['baseline_to_next_block']={**check_close(spacing['observed_baseline_to_next_block_pt'],distance,args.gap_tolerance),
                                                      'next_block':following[0]['id'] if following else None}
                if checks:result['vertical']=checks;vertical.extend(checks.values())
            answer_lines=[]
            for page,spec in sorted(page_spacing.items()):
                if not 1<=page<=actual_pages:
                    answer_lines.append({'page':page,'expected_count':spec['count'],'count_ok':None});vertical.append({'ok':None});continue
                ys=[y for text,y in page_rows(page) if len(text.replace(' ',''))>=10 and set(text.replace(' ',''))<={'_'}]
                item={'page':page,'expected_count':spec['count'],'rendered_count':len(ys),'count_ok':len(ys)==spec['count']}
                vertical.append({'ok':item['count_ok']})
                if spec['pitch'] is not None:
                    item['pitch']=check_pitch(spec['pitch'],[b-a for a,b in zip(ys,ys[1:])],args.pitch_tolerance)
                    vertical.append(item['pitch'])
                answer_lines.append(item)
        vertical_verdict=('not_requested' if not vertical else False if any(v['ok'] is False for v in vertical)
                          else None if any(v['ok'] is None for v in vertical) else True)
        bad=[r for r in results if r['line_count_matches'] is False]
        unresolved=[r for r in results if r['line_count_matches'] is None]
        pages_match=page_count==actual_pages if page_count is not None else None
        complete=bool(results) and explicit and verified and not unresolved
        line_verdict=False if bad or pages_match is False else True if complete else None
        verdict=(line_verdict if vertical_verdict in ('not_requested',True)
                 else False if vertical_verdict is False or line_verdict is False else None)
        report={'schema_version':4,'reference_verified':verified,'paragraph_boundaries_explicit':explicit,
                'blocks_checked':len(results),'matched_blocks':sum(r['status']=='matched' for r in results),
                'physical_page_count_expected':page_count,'physical_page_count_rendered':actual_pages,
                'physical_page_count_matches':pages_match,'line_count_preserved':line_verdict,
                'vertical_spacing_preserved':vertical_verdict,'all_checks_passed':verdict,
                'vertical_mismatches':[{'id':r['id'],**{k:v for k,v in r['vertical'].items() if v['ok'] is not True}}
                                       for r in results if any(v['ok'] is not True for v in r.get('vertical',{}).values())],
                'answer_lines':answer_lines,
                'exact_line_endings_blocks':sum(r.get('exact_line_endings',False) for r in results),
                'unresolved_blocks':unresolved,'line_count_mismatches':bad,'blocks':results,
                'scope':'Only listed paragraph blocks and declared spacing measurements were checked. Scan coverage, glyphs, diagrams, marks, page labels, and visual alignment need separate checks.'}
        if args.json:
            with args.json.open('x',encoding='utf-8') as stream:
                stream.write(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
        print(json.dumps(report,ensure_ascii=True,indent=2))
        return 2 if verdict is False else 0 if verdict is True else 3
    except (OSError,ValueError,KeyError,TypeError,ImportError,csv.Error) as error:
        print('Line check failed: '+json.dumps(str(error),ensure_ascii=True),file=sys.stderr);return 1


if __name__=='__main__':
    raise SystemExit(main())