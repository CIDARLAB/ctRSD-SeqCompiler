
# -*- coding: utf-8 -*-
"""

@author: tnm12

PATCH NOTES FROM VERSION 2.0.0

    
"""

import re
import xlrd
from openpyxl import Workbook, load_workbook
# xlrd.xlsx.ensure_elementtree_imported(False, None)
# xlrd.xlsx.Element_has_iter = True


'''
###############################################################################
Helper functions for ctRSD_seq_compile()
###############################################################################
'''

def rc_seq(seq, spec_out='rc', rna=0):
    # Helper function for ctRSD_seq_compile()

    new_seq = '3'+seq+'5'
    new_seq = list(new_seq)
    seq = list(seq)
    ds = ''

    for i in range(len(seq)):
        if seq[i].lower() == 'u' and rna == 0:
            seq[i] = 'T'
        elif seq[i].lower() == 't' and rna == 1:
            seq[i] = 'U'
        if seq[i].lower() == 'g':
            new_seq[i+1] = 'C'
        elif seq[i].lower() == 'c':
            new_seq[i+1] = 'G'
        elif seq[i].lower() == 't':
            new_seq[i+1] = 'A'
        elif seq[i].lower() == 'u':
            new_seq[i+1] = 'A'
        elif seq[i].lower() == 'a' and rna == 0:
            new_seq[i+1] = 'T'
        elif seq[i].lower() == 'a' and rna == 1:
            new_seq[i+1] = 'U'

    if spec_out.lower() == 'rc':
        out_seq = ds.join(new_seq[::-1])
    elif spec_out.lower() == 'c':
        out_seq = ds.join(new_seq)
    elif spec_out.lower() == 'r':
        out_seq = '3'+ds.join(seq[::-1])+'5'
    elif spec_out.lower() == 's':
        out_seq = '5'+ds.join(seq[:])+'3'

    return out_seq


def xlsheet_to_dict(filepath, sheet_name, ncol):
    # Helper function for ctRSD_seq_compile()

    if isinstance(filepath, bytes):
        seq_workbook = xlrd.open_workbook(file_contents=filepath)
    else:
        seq_workbook = xlrd.open_workbook(filepath)

    sheet = seq_workbook.sheet_by_name(sheet_name)
    sheet_dict = {}
    dv = [[0]*ncol for i in range(0, sheet.nrows)]
    for ri in range(1, sheet.nrows):
        dk = sheet.cell_value(ri, 0)
        a = 0
        for ci in range(0, ncol):
            dv[ri][a] = sheet.cell_value(ri, ci+1)
            a += 1

        sheet_dict.update({dk: dv[ri]})

    return sheet_dict

def search_xldict(sheet_dict, key):
    # Helper function for ctRSD_seq_compile()
    
    if key in sheet_dict:
        return sheet_dict[key]
    else:
        return None

def create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len):
    # Helper function for ctRSD_seq_compile()

    # full input sequence
    template = us_S + prom_S + rc_seq(ctRSD_part, spec_out='s')[1:-1] + ds_S
    # truncated sequence
    if temp_len != 0:
        # truncating template based on excess sequence downstream
        if len(template) >= temp_len and ds_S != '':
            template = template[0:temp_len]
        # truncating sequence based on excess sequence upstream
        elif len(template) >= temp_len and ds_S == '':
            strt = len(template) - temp_len
            template = template[strt:] 
        else:
            print('Issue processing template length request')

    # RNA Template
    rna_template = rc_seq(ctRSD_part, spec_out='s',rna=1)[1:-1]
        
    return template, rna_template



def domain_seq_selection(name,inps,outps,fuels,dashcheckInps,dashcheckOuts,EcheckInps,EcheckOuts,dchecks,dashIndf,RcheckOuts,r_k,\
                         eI,eO,e_d,r_d,d,d_d,s,s_d,cp,cp_d,c,c_d,CDS,cds_d,n,n_d,rflap,rflap_d,us,us_d,ds,ds_d,\
                         rna,o_bm,O_BM_d,i_bm,I_BM_d,i_th,o_th,th_d,hp5,hp5_d,Rz,Rz_d,term,term_d,prom,prom_d,L,L_d):
    
    # Helper function for ctRSD_seq_compile()
    
            
    if EcheckInps and EcheckOuts:
        ei_S = e_d[eI][0]
        ei_Sp = rc_seq(e_d[eI][0])[1:-1]
        eo_S = e_d[eO][0] 
    
    elif EcheckInps:
        ei_S = e_d[eI][0]
        ei_Sp = rc_seq(e_d[eI][0])[1:-1]
        eo_S = '' 
    
    elif EcheckOuts:
        ei_S = ''
        ei_Sp = ''
        eo_S = e_d[eO][0]
        
    else:
        ei_S = ''
        ei_Sp = ''
        eo_S = ''    
    
    if RcheckOuts:
        r_S = r_d[r_k][0]
    else:
        r_S = ''
        
        
    if dchecks:
        d_S = d_d[d][0]
    else:
        d_S = ''
        
        
    if len(s) == 0:
        s_S = ''
        s_Sp = ''
    else:
        s_S = s_d[s][0]
        s_Sp = rc_seq(s_d[s][0])[1:-1]
        
    if len(cp) == 0:
        cp_S = ''
    else:
        cp_S = cp_d[cp][0]
        
    if len(n) == 0:
        n_S = ''
    else:
        n_S = n_d[n][0]
    
    if c == 0:
        c_S = ''
    else:
        if len(cp) == 0:
            print('c domain included without complementary cp domain defined')
            #return('Function terminated. See error.')
            c_S = 'Error'
        
        elif c > len(cp_S):
            print('Length of specified c domain exceeds length of the defined complementary cp domain')
            #return('Function terminated. See error.')
            c_S = 'Error'
            
        else:
            c_S = rc_seq(cp_S[0:c])[1:-1]
    
    if len(CDS) == 0:
        cds_S = ''
    else:
        cds_S = cds_d[CDS][0]
        
    if len(rflap) == 0:
        rflap_S = ''
    else:
        rflap_S = rflap_d[rflap][0]
        
    if len(us) == 0:
        us_S = ''
    else:
        us_S = ''
        for m in range(len(us)):
            us_S = us_S + us_d[us[m]][0]
            
    if len(ds) == 0:
        ds_S = ''
    else:
        ds_S = ''
        for m in range(len(ds)):
            ds_S = ds_S + ds_d[ds[m]][0]
    
    #Extracted sequences
    if o_bm != '':    
        o_bm_S = O_BM_d[o_bm][0]
    else:
        o_bm_S = ''
    if o_th != '':
        o_th_Sp = rc_seq(th_d[o_th][0])[1:-1]
        o_th_S = th_d[o_th][0]
    else:
        o_th_Sp = ''
        o_th_S = ''
    if i_bm != '':
        i_bm_S = I_BM_d[i_bm][0]
        i_bm_Sp = rc_seq(O_BM_d[i_bm][0])[1:-1]
    else:
        i_bm_S = ''
        i_bm_Sp = ''
    if i_th != '':
        i_th_S = th_d[i_th][0]
    else:
        i_th_S = ''
    
    if inps == 1 or outps == 1 or fuels == 1:
        l_S = ''
        rz_S = ''
    else:
        l_S = L_d[L][0]
        rz_S = Rz_d[Rz][0]
        
    
    # from optional inputs
    hp5_S = hp5_d[hp5][0]
    term_S = term_d[term][0]
    prom_S = prom_d[prom][0]
    
    
    #Adding an indication of the Rz cleavage site for RNAs
    if rna == 1 and Rz[0].lower() != 'x':
        cl = '|'
    else:
        cl =''

    
    return [ei_S,ei_Sp,eo_S,r_S,d_S,s_S,s_Sp,cp_S,n_S,c_S,cds_S,rflap_S,o_bm_S,o_th_Sp,i_bm_S,i_th_S,i_bm_Sp,o_th_S,hp5_S,l_S,rz_S,term_S,prom_S,cl,us_S,ds_S]

def build_annotations(features, domain_to_seq_to_partID_d, temp_len=0):
    annotations = []
    start = 1
    for feature_map in features:
        for feature_name, feature_value in feature_map.items():
            part_type = feature_name
            part_seq = feature_value[0] if feature_value[0] else ""
            part_len = len(part_seq)


            part_id = ""
            if not part_seq or part_seq == "|":
                continue # Skip if part_seq is empty or cleavage site indicator
            elif part_seq in domain_to_seq_to_partID_d[part_type]:
                part_id = domain_to_seq_to_partID_d[part_type][part_seq]
                part_orientation = "forward"
            elif rc_seq(part_seq, spec_out='rc')[1:-1] in domain_to_seq_to_partID_d[part_type]:
                part_id = domain_to_seq_to_partID_d[part_type][rc_seq(part_seq, spec_out='rc', rna=0)[1:-1]]
                part_orientation = "reverse"
            elif len(feature_value) > 1:
                part_id = feature_value[1]
                part_orientation = "forward" # default orientation if not specified
            else:
                raise ValueError(f"Sequence '{part_seq}' not found in '{part_type}', {domain_to_seq_to_partID_d[part_type]} excel dictionary.")

            if part_seq and part_seq != "|":
                end = start + part_len - 1 if temp_len == 0 else min(start + part_len - 1, temp_len)
                annotations.append({
                    "name": "misc_feature",
                    "start": start,
                    "end": end,

                    "qualifiers": {
                        "label": part_type + "-" + part_id,
                        "note": part_id,
                        },

                    "orientation": part_orientation,
                    "part_type": part_type
                })
                start = start + part_len

    return annotations

def write_genbank_text(sequence, locus_name="ctRSD_sequence", annotations=None, rna=False):
    genbank_text_list = []

    seq_type = "RNA" if rna else "DNA"

    seq = "".join(ch for ch in str(sequence).upper() if ch.isalnum())
    if not seq:
        raise ValueError("Sequence is empty")

    genbank_text_list.append(f"LOCUS       {locus_name[:16]:<16} {len(seq)} bp    {seq_type}     linear   UNK 01-JAN-2000\n")
    genbank_text_list.append(f"DEFINITION  {locus_name}\n")
    genbank_text_list.append("ACCESSION   \n")
    genbank_text_list.append("VERSION     \n")
    genbank_text_list.append("KEYWORDS    .\n")
    genbank_text_list.append("SOURCE      synthetic construct\n")
    genbank_text_list.append("  ORGANISM  synthetic construct\n")
    genbank_text_list.append("FEATURES             Location/Qualifiers\n")
    genbank_text_list.append(f"     source          1..{len(seq)}\n")
    genbank_text_list.append("                     /organism=\"synthetic construct\"\n")
    genbank_text_list.append(f"                     /mol_type=\"other {seq_type}\"\n")

    if annotations:
        for ann in annotations:
            if isinstance(ann, dict):
                feature_name = ann.get("name", "feature")
                start = ann.get("start", 1)
                end = ann.get("end", len(seq))
                qualifiers = ann.get("qualifiers", {})
                orientation = ann.get("orientation", None)
            else:
                feature_name, start, end, qualifiers, orientation = ann

            genbank_text_list.append(f"     {feature_name:<16} {f'complement(' if orientation == 'reverse' else ''}{start}..{end}{')' if orientation == 'reverse' else ''}\n")
            if isinstance(qualifiers, dict):
                for key, value in qualifiers.items():
                    if value is not None:
                        genbank_text_list.append(f'                     /{key}="{value}"\n')
            elif isinstance(qualifiers, (list, tuple)):
                for qual in qualifiers:
                    genbank_text_list.append(f"                     {qual}\n")

    genbank_text_list.append("ORIGIN\n")
    for start in range(0, len(seq), 60):
        chunk = seq[start:start+60]
        genbank_text_list.append(f"{start+1:>9} ")
        genbank_text_list.append(" ".join(chunk[i:i+10] for i in range(0, len(chunk), 10)))
        genbank_text_list.append("\n")
    genbank_text_list.append("//\n")

    return "".join(genbank_text_list)

def domain_dict_to_seq_dict(domain_dict, suffix=""):
    return {
        # Reverse Complement sequences
        **{rc_seq(value[0], spec_out='rc')[1:-1]: key + suffix + "_rc" for key, value in domain_dict.items()},

        # Forward sequences (overrides any duplicates with the same sequence)
        **{value[0]: key + suffix for key, value in domain_dict.items()}
    }


def reverse_dict(d):
    return {v: k for k, v in d.items()}


'''
###############################################################################
Class definition
###############################################################################
'''

class RSD_sim:
    def __init__(self):
        self.N = 0  # Initialize the number of species to 0        

    def load_domains(self, filepath):
        '''
        This function compiles the domain sequences from the excel file into a dictionary
        that can be used to generate sequences for the ctRSD simulator.
        '''

        # Key to Sequence Dictionaries
        self.th_d = xlsheet_to_dict(filepath, 'th', 1)  # toehold sequences (th')
        self.I_BM_d = xlsheet_to_dict(filepath, 'I BM', 1) # input branch migration seqeuences
        self.O_BM_d = xlsheet_to_dict(filepath, 'O BM', 1) # output branch migration seqeuences
        self.r_d = xlsheet_to_dict(filepath, 'r', 1)  # reversible domain on reporters
        self.L_d = xlsheet_to_dict(filepath, 'L', 1)  # Linker for ribozyme
        self.e_d = xlsheet_to_dict(filepath, 'e', 1) # extended branch migration domain
        self.Rz_d = xlsheet_to_dict(filepath, 'Rz', 1)  # ribozymes
        self.term_d = xlsheet_to_dict(filepath, 'terminators', 1)  # terminators
        self.prom_d = xlsheet_to_dict(filepath, 'promoters', 1)  # promoters
        self.hp5_d = xlsheet_to_dict(filepath, '5hp', 1)  # 5' hairpin sequences
        self.s_d = xlsheet_to_dict(filepath, 's', 1)  # spacer sequences
        self.us_d = xlsheet_to_dict(filepath, 'US', 1) # sequences upstream of promoter
        self.ds_d = xlsheet_to_dict(filepath, 'DS', 1) # sequences downstream of terminator
        self.cds_d = xlsheet_to_dict(filepath, 'CDS', 1) # CDS sequences for protein outputs
        self.rflap_d = xlsheet_to_dict(filepath, 'rflap', 1) # fluorescent RNA aptamer sequences
        self.cp_d = xlsheet_to_dict(filepath, 'c prime', 1) # c' linker between gate and n
        self.n_d = xlsheet_to_dict(filepath, 'n', 1) # N-terminal linker between c' and CDS
        self.c_d = xlsheet_to_dict(filepath, 'c', 1) # clamp adjacent to toehold, complementary to cp
        self.d_d = xlsheet_to_dict(filepath, 'd', 1) # distal domain for translation

        self.domain_to_name_d = {
            "th": "http://identifiers.org/so/SO:0001263/toehold",
            "i_bm": "http://identifiers.org/so/SO:0001263/inputBranchMigration",
            "o_bm": "http://identifiers.org/so/SO:0001263/outputBranchMigration",
            "r": "http://identifiers.org/so/SO:0001263/outputReportingDomain",
            "l": "http://identifiers.org/so/SO:0001263/linker",
            "e": "http://identifiers.org/so/SO:0001263/extendedBranchMigration",
            "rz": "http://identifiers.org/so/SO:0001263/ribozyme",
            "term": "http://identifiers.org/so/SO:0000141", # terminator
            "prom": "http://identifiers.org/so/SO:0000167", # promoter
            "hp5": "http://identifiers.org/so/SO:0001263/hp5",
            "s": "http://identifiers.org/so/SO:0002223", # spacer
            "us": "us", # TODO: upstream sequence, add appropriate SO identifier
            "ds": "ds", # TODO: downstream sequence, add appropriate SO identifier
            "cds": "http://identifiers.org/so/SO:0000316", # cds
            "rflap": "http://identifiers.org/so/SO:0001263/rflap",
            "cp": "http://identifiers.org/so/SO:0001263/clampDomain",    # Clamp Domain Prime points to 'c' sheet
            "n": "http://identifiers.org/so/SO:0001263/nTerminalLinker",
            "c": "http://identifiers.org/so/SO:0001263/clampDomain",
            "d": "http://identifiers.org/so/SO:0001263/distalDomain",
            "cl": "http://identifiers.org/so/SO:0001977" # Ribonuclease cleavage site
        }

        # {domain_name : {sequence : partID}}, assumes that sequences are unique across partIDs for a given domain
        self.domain_to_seq_to_partID_d = {
            self.domain_to_name_d["th"]: domain_dict_to_seq_dict(self.th_d),
            self.domain_to_name_d["i_bm"]: domain_dict_to_seq_dict(self.I_BM_d, "_ibm"),
            self.domain_to_name_d["o_bm"]: domain_dict_to_seq_dict(self.O_BM_d, "_obm"),
            self.domain_to_name_d["r"]: domain_dict_to_seq_dict(self.r_d),
            self.domain_to_name_d["l"]: domain_dict_to_seq_dict(self.L_d),
            self.domain_to_name_d["e"]: domain_dict_to_seq_dict(self.e_d),
            self.domain_to_name_d["rz"]: domain_dict_to_seq_dict(self.Rz_d),
            self.domain_to_name_d["term"]: domain_dict_to_seq_dict(self.term_d),
            self.domain_to_name_d["prom"]: domain_dict_to_seq_dict(self.prom_d),
            self.domain_to_name_d["hp5"]: domain_dict_to_seq_dict(self.hp5_d),
            self.domain_to_name_d["s"]: domain_dict_to_seq_dict(self.s_d),
            self.domain_to_name_d["us"]: domain_dict_to_seq_dict(self.us_d),
            self.domain_to_name_d["ds"]: domain_dict_to_seq_dict(self.ds_d),
            self.domain_to_name_d["cds"]: domain_dict_to_seq_dict(self.cds_d),
            self.domain_to_name_d["rflap"]: domain_dict_to_seq_dict(self.rflap_d),
            #self.domain_to_name_d["cp"]: domain_dict_to_seq_dict(self.cp_d),
            self.domain_to_name_d["n"]: domain_dict_to_seq_dict(self.n_d),
            self.domain_to_name_d["c"]: domain_dict_to_seq_dict(self.c_d),
            self.domain_to_name_d["d"]: domain_dict_to_seq_dict(self.d_d)
        }

        # {domain_name : {partID : sequence}}, assumes that sequences are unique across partIDs for a given domain
        self.domain_to_partID_to_seq_d = {k: reverse_dict(v) for k, v in self.domain_to_seq_to_partID_d.items()}

        #print("Domain to sequence to part ID mapping:", self.domain_to_seq_to_partID_d[self.domain_to_name_d["c"]])
           
            
    def ctRSD_seq_compile(self,
        name,
        Rz='Ro',
        L='L',
        term='T7t',
        hp5='5hp',
        prom='T7p',
        eI='',
        eO='',
        s='',
        invert=0,
        invL='A',
        agL='TA',
        AGiloop=5,
        otype=1,
        rna=0,
        us=[],
        ds=[],
        temp_len=0,
        cp='',
        n='',
        c=0,
        d='',
        CDS='',
        rflap='',
        get_genbank=1):
        
        '''
        TO ADD:
            AND gate output generator? 
            Custom sequence input options for some of the optional variables
                If the key provided doesnt match then assume it is custom seq to insert
            
        '''
        
        th_d = self.th_d
        I_BM_d = self.I_BM_d
        O_BM_d = self.O_BM_d
        r_d = self.r_d
        L_d = self.L_d
        e_d = self.e_d
        Rz_d = self.Rz_d
        term_d = self.term_d
        prom_d = self.prom_d
        hp5_d = self.hp5_d
        s_d = self.s_d
        us_d = self.us_d
        ds_d = self.ds_d
        cds_d = self.cds_d
        rflap_d = self.rflap_d
        cp_d = self.cp_d
        n_d = self.n_d
        c_d = self.c_d
        d_d = self.d_d
        domain_to_name_d = self.domain_to_name_d
        domain_to_seq_to_partID_d = self.domain_to_seq_to_partID_d
        
        # converting into th' as code was originally written to handle that
        for m in th_d.keys():
            th_d[m] = [rc_seq(th_d[m][0])[1:-1]]

        i_th = ''
        o_th = ''
        i_bm = ''
        o_bm = ''
        
        #gate = re.compile(r"(g|gate)\{\w\d+(\-\d+)*e*\,\w\d+(\-\d+)*(r*|e*)\}")
        gate = re.compile(r"(g|gate)\{\w+\d+(\-\d+)*e*\,\w+\d+(\-\d+)*(r*|e*|d*|ed*)\}")
        gates = gate.fullmatch(name.lower())
        #print(gates)        
        
        AG = re.compile(r"(ag|g|gate)\{\w+\d+(\-\d+)*(\.|\&)\w+\d+(\-\d+)*e*\,\w+\d+(\-\d+)*(r*|e*|d*|ed*)\}")
        AGs = AG.fullmatch(name.lower())
        
        inp = re.compile(r"(i|in|inp|input)\{\w+\d+(\-\d+)*(r*|e*|d*|ed*)\}")
        inps = inp.fullmatch(name.lower())
        
        outp = re.compile(r"(o|out|output)\{\w*\d+(\-\d+)*e*\,\w+\d+(\-\d+)*(r*|e*|d*|ed*)\}")
        outps = outp.fullmatch(name.lower())
        
        thresh = re.compile(r"(tg|t|th)\{\w+\d+(\-\d+)*(r*|e*|d*|ed*)\}")
        threshs = thresh.fullmatch(name.lower())
        
        fuel = re.compile(r"f\{\w+\d+(\-\d+)*e*\}")
        fuels = fuel.fullmatch(name.lower())
        
        CG = re.compile(r"cg\{\w+\d+(\-\d+)*\,\w+\d+(\-\d+)*\}")
        CGs = CG.fullmatch(name.lower())
        
        
        # extension domains on inputs or outputs
        if threshs:
           EcheckInp = re.compile(r"e")
           EcheckInps = EcheckInp.search(name.lower())
           EcheckOuts = 0
        else:
            EcheckInp = re.compile(r"e\,")
            EcheckInps = EcheckInp.search(name.lower())

            
            EcheckOut = re.compile(r"e[^,]")
            EcheckOuts = EcheckOut.search(name.lower())

            
        if (eI == '' and EcheckInps) or (eO == '' and EcheckOuts):
            print('e included in species name, but no domain was provided in function declaration.')
            return('Function terminated, see error.')
        elif (eI != '' and EcheckInps == None) or (eO != '' and EcheckOuts == None):
            print('e included as input but not in name, the domain will not be in the sequence.')
            return('Function terminated, see error.')
        
        
        # distal domains
        dcheck = re.compile(r"d\}")
        dchecks =dcheck.search(name.lower())   

        
        if (d == '' and dchecks):
            print('d included in species name, but no domain number was provided in function declaration.')
            return('Function terminated, see error.')
        elif (d != '' and dchecks == None):
            print('d included as input but not in name, the domain will not be in the sequence.')
            return('Function terminated, see error.')
        
        
        # checking for dashes for longer named toeholds
        dashcheckInp = re.compile(r"\-\d+e*\,")
        dashcheckInps = dashcheckInp.search(name.lower())

        dashcheckOut = re.compile(r"\-\d+(r*|e*|d*)\}")
        dashcheckOuts = dashcheckOut.search(name.lower())

        dashIndf = re.findall(r"\-\d+",name.lower())

        if gates:
            Indf = re.findall(r"[a-z]+\d+",name.lower())
        elif outps:
            Indf = re.findall(r"[a-z]*\d+",name.lower())
        elif inps or threshs or fuels or CGs: 
            Indf = re.findall(r"[a-z]\d+",name.lower())
            
        
        # checking for the r domain (reversible reporter domain, or thresholding domain)
        RcheckOut = re.compile(r"r\}")
        RcheckOuts = RcheckOut.search(name.lower())
        
        if gates:
            '''
            ###################################################################
            GATES
            ###################################################################
            '''
            
            if dashcheckInps and dashcheckOuts:
                i_th = Indf[0]
                o_th = Indf[1]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+dashIndf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[1][1:]
                else:
                    r_k = ''
                
            elif dashcheckInps:
                i_th = Indf[0]
                o_th = Indf[1][0]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+Indf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[1][1:]
                else:
                    r_k = ''
                
            elif dashcheckOuts:
                i_th = Indf[0][0]
                o_th = Indf[1]
                i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+dashIndf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[0][1:]
                else:
                    r_k = ''
                    
            else:
                i_th = Indf[0][0]
                o_th = Indf[1][0]
                i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+Indf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[1][1:]
                else:
                    r_k = ''
                    
            ###################################################################
            # Selecting the correct seqeuences for each domain
            ###################################################################
            [ei_S,ei_Sp,eo_S,r_S,d_S,s_S,s_Sp,cp_S,n_S,c_S,cds_S,rflap_S,o_bm_S,o_th_Sp,i_bm_S,i_th_S,i_bm_Sp,o_th_S,hp5_S,l_S,rz_S,term_S,prom_S,cl,us_S,ds_S] = \
            domain_seq_selection(name,inps,outps,fuels,dashcheckInps,dashcheckOuts,EcheckInps,EcheckOuts,dchecks,dashIndf,RcheckOuts,r_k,\
                                     eI,eO,e_d,r_d,d,d_d,s,s_d,cp,cp_d,c,c_d,CDS,cds_d,n,n_d,rflap,rflap_d,us,us_d,ds,ds_d,\
                                     rna,o_bm,O_BM_d,i_bm,I_BM_d,i_th,o_th,th_d,hp5,hp5_d,Rz,Rz_d,term,term_d,prom,prom_d,L,L_d)

            if c_S == 'Error':
                return('Function terminated, see error.')
                
            ###################################################################
            # Stitching gate sequences
            ###################################################################
            
            if invert == 0:
            
                ctRSD_part = hp5_S +\
                             d_S + r_S + eo_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S +\
                             l_S + cl + rz_S +\
                             s_S + i_th_S + i_bm_Sp + ei_Sp + o_th_S +\
                             cp_S + n_S + cds_S + rflap_S +\
                             term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [eo_S]},
                    {domain_to_name_d["o_bm"]: [o_bm_S]},
                    {domain_to_name_d["c"]: [c_S]},
                    {domain_to_name_d["th"]: [o_th_Sp]},
                    {domain_to_name_d["e"] : [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["s"]: [s_S]},
                    {domain_to_name_d["th"]: [i_th_S]},
                    {domain_to_name_d["o_bm"]: [i_bm_Sp]},
                    {domain_to_name_d["e"]: [ei_Sp]},
                    {domain_to_name_d["th"]: [o_th_S]},
                    {domain_to_name_d["cp"]: [cp_S]},
                    {domain_to_name_d["n"]: [n_S]},
                    {domain_to_name_d["cds"]: [cds_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]

                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                
            elif invert == 1:
                    
                ctRSD_part = hp5_S +\
                             s_S + i_th_S + i_bm_Sp + ei_Sp + o_th_S +\
                             cp_S + n_S + cds_S + rflap_S +\
                             l_S + cl + rz_S + invL +\
                             d_S + r_S + eo_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S +\
                             term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["s"]: [s_S]},
                    {domain_to_name_d["th"]: [i_th_S]},
                    {domain_to_name_d["o_bm"]: [i_bm_Sp]},
                    {domain_to_name_d["e"]: [ei_Sp]},
                    {domain_to_name_d["th"]: [o_th_S]},
                    {domain_to_name_d["cp"]: [cp_S]},
                    {domain_to_name_d["n"]: [n_S]},
                    {domain_to_name_d["cds"]: [cds_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["l"]: [invL]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [eo_S]},
                    {domain_to_name_d["o_bm"]: [o_bm_S]},
                    {domain_to_name_d["c"]: [c_S]},
                    {domain_to_name_d["th"]: [o_th_Sp]},
                    {domain_to_name_d["e"]: [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]
                
                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)

        
        elif outps:
            '''
            ###################################################################
            OUTPUTS
            ###################################################################
            '''
            
            if dashcheckInps and dashcheckOuts:
                o_th = Indf[2]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+dashIndf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[-1][1:]
                else:
                    r_k = ''
                
            elif dashcheckInps:
                o_th = Indf[-1][0]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+Indf[-1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[-1][1:]
                else:
                    r_k = ''
                
            elif dashcheckOuts:
                o_th = Indf[1]
                if len(re.findall(r"[a-z]\d+",Indf[0].lower()))==0:
                    i_bm = 'D'+Indf[0][:]
                else:
                    i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+dashIndf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[-1][1:]
                else:
                    r_k = ''
                    
            else:
                o_th = Indf[1][0]
                if len(re.findall(r"[a-z]\d+",Indf[0].lower()))==0:
                    i_bm = 'D'+Indf[0][:]
                else:
                    i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+Indf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[-1][1:]
                else:
                    r_k = ''
            
            
            


            ###################################################################
            # Selecting the correct seqeuences for each domain
            ###################################################################
            [ei_S,ei_Sp,eo_S,r_S,d_S,s_S,s_Sp,cp_S,n_S,c_S,cds_S,rflap_S,o_bm_S,o_th_Sp,i_bm_S,i_th_S,i_bm_Sp,o_th_S,hp5_S,l_S,rz_S,term_S,prom_S,cl,us_S,ds_S] = \
            domain_seq_selection(name,inps,outps,fuels,dashcheckInps,dashcheckOuts,EcheckInps,EcheckOuts,dchecks,dashIndf,RcheckOuts,r_k,\
                                     eI,eO,e_d,r_d,d,d_d,s,s_d,cp,cp_d,c,c_d,CDS,cds_d,n,n_d,rflap,rflap_d,us,us_d,ds,ds_d,\
                                     rna,o_bm,O_BM_d,i_bm,I_BM_d,i_th,o_th,th_d,hp5,hp5_d,Rz,Rz_d,term,term_d,prom,prom_d,L,L_d)
            
            
            if otype == 1:
                rz_S = Rz_d[Rz][0]
            else:
                rz_S = ''
                
            ###################################################################
            # Stitching output sequences
            ###################################################################
            
            if invert == 0:
            
                ctRSD_part = hp5_S +\
                             d_S + r_S + eo_S + o_bm_S + o_th_Sp + s_Sp + ei_S + i_bm_S +\
                             l_S + cl + rz_S +\
                             rflap_S + term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [eo_S]},
                    {domain_to_name_d["o_bm"]: [o_bm_S]},
                    {domain_to_name_d["th"]: [o_th_Sp]},
                    {domain_to_name_d["s"]: [s_Sp]},
                    {domain_to_name_d["e"]: [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]
                
                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                
            elif invert == 1:
                    
                ctRSD_part = hp5_S +\
                             l_S + cl + rz_S + invL +\
                             d_S + r_S + eo_S + o_bm_S + o_th_Sp + s_Sp + ei_S + i_bm_S +\
                             rflap_S + term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["l"]: [invL]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [eo_S]},
                    {domain_to_name_d["o_bm"]: [o_bm_S]},
                    {domain_to_name_d["th"]: [o_th_Sp]},
                    {domain_to_name_d["s"]: [s_Sp]},
                    {domain_to_name_d["e"]: [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]
                
                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                
                
        elif inps:
            '''
            ###################################################################
            INPUTS
            ###################################################################
            '''
            
            if dashcheckOuts:
                o_th = Indf[0]
                o_bm = 'D'+dashIndf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[0][1:]
                else:
                    r_k = ''
            else:
                o_th = Indf[0][0]
                o_bm = 'D'+Indf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[0][1:]
                else:
                    r_k = ''
                    
            
            ###################################################################
            # Selecting the correct seqeuences for each domain
            ###################################################################
            [ei_S,ei_Sp,eo_S,r_S,d_S,s_S,s_Sp,cp_S,n_S,c_S,cds_S,rflap_S,o_bm_S,o_th_Sp,i_bm_S,i_th_S,i_bm_Sp,o_th_S,hp5_S,l_S,rz_S,term_S,prom_S,cl,us_S,ds_S] = \
            domain_seq_selection(name,inps,outps,fuels,dashcheckInps,dashcheckOuts,EcheckInps,EcheckOuts,dchecks,dashIndf,RcheckOuts,r_k,\
                                     eI,eO,e_d,r_d,d,d_d,s,s_d,cp,cp_d,c,c_d,CDS,cds_d,n,n_d,rflap,rflap_d,us,us_d,ds,ds_d,\
                                     rna,o_bm,O_BM_d,i_bm,I_BM_d,i_th,o_th,th_d,hp5,hp5_d,Rz,Rz_d,term,term_d,prom,prom_d,L,L_d)
            
            
            ###################################################################
            # Stitching input sequences
            ###################################################################
            
            ctRSD_part = hp5_S + d_S + r_S + eo_S + o_bm_S + o_th_Sp + s_Sp + rflap_S + term_S

            rna_features = [
                {domain_to_name_d["hp5"]: [hp5_S]},
                {domain_to_name_d["d"]: [d_S]},
                {domain_to_name_d["r"]: [r_S]},
                {domain_to_name_d["e"]: [eo_S]},
                {domain_to_name_d["o_bm"]: [o_bm_S]},
                {domain_to_name_d["th"]: [o_th_Sp]},
                {domain_to_name_d["s"]: [s_Sp]},
                {domain_to_name_d["rflap"]: [rflap_S]},
                {domain_to_name_d["term"]: [term_S]}
            ]
            
            template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
            
        
        elif threshs:
            '''
            ###################################################################
            THRESHOLD GATES
            ###################################################################
            '''
            
            if dashcheckInps:
                i_th = Indf[0]
                i_bm = 'D'+dashIndf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[0][1:]
                else:
                    r_k = ''
                
            else:
                i_th = Indf[0][0]
                i_bm = 'D'+Indf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[0][1:]
                else:
                    r_k = ''
                    
            ###################################################################
            # Selecting the correct seqeuences for each domain
            ###################################################################
            [ei_S,ei_Sp,eo_S,r_S,d_S,s_S,s_Sp,cp_S,n_S,c_S,cds_S,rflap_S,o_bm_S,o_th_Sp,i_bm_S,i_th_S,i_bm_Sp,o_th_S,hp5_S,l_S,rz_S,term_S,prom_S,cl,us_S,ds_S] = \
            domain_seq_selection(name,inps,outps,fuels,dashcheckInps,dashcheckOuts,EcheckInps,EcheckOuts,dchecks,dashIndf,RcheckOuts,r_k,\
                                     eI,eO,e_d,r_d,d,d_d,s,s_d,cp,cp_d,c,c_d,CDS,cds_d,n,n_d,rflap,rflap_d,us,us_d,ds,ds_d,\
                                     rna,o_bm,O_BM_d,i_bm,I_BM_d,i_th,o_th,th_d,hp5,hp5_d,Rz,Rz_d,term,term_d,prom,prom_d,L,L_d)
            
            if c_S == 'Error':
                return('Function terminated, see error.')
            
            d_Sp = rc_seq(d_S)[1:-1]
            
            ###################################################################
            # Stitching thresholding gate sequences
            ###################################################################
            
            if invert == 0:
            
                ctRSD_part = hp5_S +\
                             d_S + c_S + r_S + ei_S + i_bm_S +\
                             l_S + cl + rz_S +\
                             s_S + i_th_S + i_bm_Sp + ei_Sp + rc_seq(r_S)[1:-1] +\
                             d_Sp + cp_S + n_S + cds_S + rflap_S +\
                             term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["c"]: [c_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["s"]: [s_S]},
                    {domain_to_name_d["th"]: [i_th_S]},
                    {domain_to_name_d["o_bm"]: [i_bm_Sp]},
                    {domain_to_name_d["e"]: [ei_Sp]},
                    {domain_to_name_d["r"]: [rc_seq(r_S)[1:-1]]},
                    {domain_to_name_d["d"]: [d_Sp]},
                    {domain_to_name_d["cp"]: [cp_S]},
                    {domain_to_name_d["n"]: [n_S]},
                    {domain_to_name_d["cds"]: [cds_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]
                
                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                
            elif invert == 1:
                    
                ctRSD_part = hp5_S +\
                             s_S + i_th_S + i_bm_Sp + rc_seq(r_S)[1:-1] +\
                             d_Sp + cp_S + n_S + cds_S + rflap_S +\
                             l_S + cl + rz_S + invL +\
                             d_S + c_S + r_S + ei_S + i_bm_S +\
                             term_S

                rna_features = [
                    {domain_to_name_d["hp5"]: [hp5_S]},
                    {domain_to_name_d["s"]: [s_S]},
                    {domain_to_name_d["th"]: [i_th_S]},
                    {domain_to_name_d["o_bm"]: [i_bm_Sp]},
                    {domain_to_name_d["rc_r"]: [rc_seq(r_S)[1:-1]]},
                    {domain_to_name_d["d"]: [d_Sp]},
                    {domain_to_name_d["cp"]: [cp_S]},
                    {domain_to_name_d["n"]: [n_S]},
                    {domain_to_name_d["cds"]: [cds_S]},
                    {domain_to_name_d["rflap"]: [rflap_S]},
                    {domain_to_name_d["l"]: [l_S]},
                    {domain_to_name_d["cl"]: [cl]},
                    {domain_to_name_d["rz"]: [rz_S]},
                    {domain_to_name_d["l"]: [invL]},
                    {domain_to_name_d["d"]: [d_S]},
                    {domain_to_name_d["c"]: [c_S]},
                    {domain_to_name_d["r"]: [r_S]},
                    {domain_to_name_d["e"]: [ei_S]},
                    {domain_to_name_d["i_bm"]: [i_bm_S]},
                    {domain_to_name_d["term"]: [term_S]}
                ]
                
                template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                
                
        elif fuels:
            '''
            ###################################################################
            FUELS
            ###################################################################
            '''
            
            dashcheckOut = re.compile(r"\-\d+e*\}")
            dashcheckOuts = dashcheckOut.search(name.lower())
            
            EcheckOut = re.compile(r"e\}")
            EcheckOuts = EcheckOut.search(name.lower())
            
            Indf = re.findall(r"[a-z]\d+",name.lower())
            dashIndf = re.findall(r"\-\d+",name.lower())
            
                
            if dashcheckOuts:
                o_th = Indf[0]
                o_bm = 'D'+dashIndf[0][1:]
                
            else:
                o_th = Indf[0][0]
                o_bm = 'D'+Indf[0][1:]
                    
            if EcheckOuts:
                eo_S = e_d[eO][0]
            else: 
                eo_S = ''
                
            if len(us) == 0:
                us_S = ''
            else:
                us_S = ''
                for m in range(len(us)):
                    us_S = us_S + us_d[us[m]][0]
                    
            if len(ds) == 0:
                ds_S = ''
            else:
                ds_S = ''
                for m in range(len(ds)):
                    ds_S = ds_S + ds_d[ds[m]][0]
                
            if len(rflap) == 0:
                rflap_S = ''
            else:
                rflap_S = rflap_d[rflap][0]
                
            # Extracted sequence domains
            o_bm_S = O_BM_d[o_bm][0]
            o_th_Sp = rc_seq(th_d[o_th][0])[1:-1]
            # from optional inputs
            hp5_S = hp5_d[hp5][0]
            term_S = term_d[term][0]
            prom_S = prom_d[prom][0]
            
            ###################################################################
            # Stitching fuel sequences
            ###################################################################
            
            ctRSD_part = hp5_S + o_th_Sp + eo_S + o_bm_S + rflap_S + term_S

            rna_features = [
                {"hp5": [hp5_S]},
                {"th": [o_th_Sp]},
                {"e": [eo_S]},
                {"o_bm": [o_bm_S]},
                {"rflap": [rflap_S]},
                {"term": [term_S]}
            ]
            
            template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)

            
        elif AGs:
            '''
            ###################################################################
            AND GATES
            ###################################################################
            '''
            
            dashcheckInp1 = re.compile(r"\-\d+(\.|\&)")
            dashcheckInps1 = dashcheckInp1.search(name.lower())
            dashcheckInp2 = re.compile(r"\-\d+\,")
            dashcheckInps2 = dashcheckInp2.search(name.lower())
            
            dashcheckOut = re.compile(r"\-\d+r*\}")
            dashcheckOuts = dashcheckOut.search(name.lower())
            
            EcheckInp = re.compile(r"e\,")
            EcheckInps = EcheckInp.search(name.lower())
            
            EcheckOut = re.compile(r"e\}")
            EcheckOuts = EcheckOut.search(name.lower())
            
            RcheckOut = re.compile(r"r\}")
            RcheckOuts = RcheckOut.search(name.lower())
            
            Indf = re.findall(r"[a-z]\d+",name.lower())
            dashIndf = re.findall(r"\-\d+",name.lower())
            
            
            if dashcheckInps1 and dashcheckInps2 and dashcheckOuts:
                i_th1 = Indf[0]
                i_th2 = Indf[1]
                o_th = Indf[2]
                i_bm1 = 'D'+dashIndf[0][1:]
                i_bm2 = 'D'+dashIndf[1][1:]
                o_bm = 'D'+dashIndf[2][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[2][1:]
                
            elif dashcheckInps1 and dashcheckInps2:
                i_th1 = Indf[0]
                i_th2 = Indf[1]
                o_th = Indf[2][0]
                i_bm1 = 'D'+dashIndf[0][1:]
                i_bm2 = 'D'+dashIndf[1][1:]
                o_bm = 'D'+Indf[2][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[2][1:]
                
            elif dashcheckInps1 and dashcheckOuts:
                i_th1 = Indf[0]
                i_th2 = Indf[1][0]
                o_th = Indf[2]
                i_bm1 = 'D'+dashIndf[0][1:]
                i_bm2 = 'D'+Indf[1][1:]
                o_bm = 'D'+dashIndf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[1][1:]
                    
            elif dashcheckInps2 and dashcheckOuts:
                i_th1 = Indf[0][0]
                i_th2 = Indf[1]
                o_th = Indf[2]
                i_bm1 = 'D'+Indf[0][1:]
                i_bm2 = 'D'+dashIndf[0][1:]
                o_bm = 'D'+dashIndf[1][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[1][1:]
            
            elif dashcheckInps1:
                i_th1 = Indf[0]
                i_th2 = Indf[1][0]
                o_th = Indf[2][0]
                i_bm1 = 'D'+dashIndf[0][1:]
                i_bm2 = 'D'+Indf[1][1:]
                o_bm = 'D'+Indf[2][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[2][1:]
                
            elif dashcheckInps2:
                i_th1 = Indf[0][0]
                i_th2 = Indf[1]
                o_th = Indf[2][0]
                i_bm1 = 'D'+Indf[0][1:]
                i_bm2 = 'D'+dashIndf[0][1:]
                o_bm = 'D'+Indf[2][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[2][1:]
            
            elif dashcheckOuts:
                i_th1 = Indf[0][0]
                i_th2 = Indf[1][0]
                o_th = Indf[2]
                i_bm1 = 'D'+Indf[0][1:]
                i_bm2 = 'D'+Indf[1][1:]
                o_bm = 'D'+dashIndf[0][1:]
                
                if RcheckOuts:
                    r_k = 'r'+dashIndf[0][1:]
                           
            else:
                i_th1 = Indf[0][0]
                i_th2 = Indf[1][0]
                o_th = Indf[2][0]
                i_bm1 = 'D'+Indf[0][1:]
                i_bm2 = 'D'+Indf[1][1:]
                o_bm = 'D'+Indf[2][1:]
                
                if RcheckOuts:
                    r_k = 'r'+Indf[2][1:]
        
            if EcheckInps and EcheckOuts:
                
                ei_S = e_d[eI][0]
                ei_Sp = rc_seq(e_d[eI][0])[1:-1]
                eo_S = e_d[eO][0] 
            
            elif EcheckInps:
                
                ei_S = e_d[eI][0]
                ei_Sp = rc_seq(e_d[eI][0])[1:-1]
                eo_S = '' 
            
            elif EcheckOuts:
                
                ei_S = ''
                ei_Sp = ''
                eo_S = e_d[eO][0]
                
            else:
                ei_S = ''
                ei_Sp = ''
                eo_S = ''  
                
            if RcheckOuts:
                r_S = r_d[r_k][0]
            else:
                r_S = ''
                
            if len(s) == 0:
                s_S = ''
            else:
                s_S = s_d[s][0]
            
            if len(cp) == 0:
                cp_S = ''
            else:
                cp_S = cp_d[cp][0]
                
            if len(n) == 0:
                n_S = ''
            else:
                n_S = n_d[n][0]
                
            
            if c == 0:
                c_S = ''
            else:
                if len(cp) == 0:
                    print('c domain included without complementary cp domain defined')
                    return('Function terminated, see error.')
                
                elif c > len(cp_S):
                    print('Length of specified c domain exceeds length of the defined complementary cp domain')
                    return('Function terminated, see error.')
                    
                else:
                    c_S = rc_seq(cp_S[0:c])[1:-1]
                    
            
            if len(CDS) == 0:
                cds_S = ''
            else:
                cds_S = cds_d[CDS][0]
                
            if len(rflap) == 0:
                rflap_S = ''
            else:
                rflap_S = rflap_d[rflap][0]
                
            if len(us) == 0:
                us_S = ''
            else:
                us_S = ''
                for n in range(len(us)):
                    us_S = us_S + us_d[us[m]][0]
                    
            if len(ds) == 0:
                ds_S = ''
            else:
                ds_S = ''
                for n in range(len(ds)):
                    ds_S = ds_S + ds_d[ds[m]][0]
                
            o_bm_S = O_BM_d[o_bm][0]
            o_th_Sp = rc_seq(th_d[o_th][0])[1:-1]
            i_bm_S1 = I_BM_d[i_bm1][0]
            i_bm_S2 = I_BM_d[i_bm2][0]
            i_th_S1 = th_d[i_th1][0]
            i_th_S2 = th_d[i_th2][0]
            i_th_Sp2 = rc_seq(th_d[i_th2][0])[1:-1]
            i_bm_Sp1 = rc_seq(O_BM_d[i_bm1][0])[1:-1]
            i_bm_Sp2 = rc_seq(O_BM_d[i_bm2][0])[1:-1]
            o_th_S = th_d[o_th][0]
             # from optional inputs
            hp5_S = hp5_d[hp5][0]
            l_S = L_d[L][0]
            rz_S = Rz_d[Rz][0]
            term_S = term_d[term][0]
            prom_S = prom_d[prom][0]
            agl_S = agL
            
            #Adding an indication of the Rz cleavage site for RNAs
            if rna == 1 and Rz[0].lower() != 'x':
                cl = '|'
            else:
                cl =''

            
            ###################################################################
            # Stitching and gate sequences
            ###################################################################
            
            if invert == 0:
            
                if AGiloop == 5: # internal loop size
                    
                    ctRSD_part = hp5_S +\
                                 eo_S + r_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S2 + agl_S + i_th_Sp2[-1] + ei_S + i_bm_S1 +\
                                 l_S + cl + rz_S +\
                                 s_S + i_th_S1 + i_bm_Sp1 + ei_Sp + i_th_S2 + i_bm_Sp2 + ei_Sp + o_th_S +\
                                 cp_S + n_S + cds_S + rflap_S +\
                                 term_S

                    rna_features = [
                        {domain_to_name_d["hp5"]: [hp5_S]},
                        {domain_to_name_d["e"]: [eo_S]},
                        {domain_to_name_d["r"]: [r_S]},
                        {domain_to_name_d["o_bm"]: [o_bm_S]},
                        {domain_to_name_d["c"]: [c_S]},
                        {domain_to_name_d["th"]: [o_th_Sp]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S2]},
                        {domain_to_name_d["agl"]: [agl_S]},
                        {domain_to_name_d["th"]: [i_th_Sp2[-1]]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S1]},
                        {domain_to_name_d["l"]: [l_S]},
                        {domain_to_name_d["cl"]: [cl]},
                        {domain_to_name_d["rz"]: [rz_S]},
                        {domain_to_name_d["spacer"]: [s_S]},
                        {domain_to_name_d["th"]: [i_th_S1]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp1]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [i_th_S2]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp2]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [o_th_S]},
                        {domain_to_name_d["cp"]: [cp_S]},
                        {domain_to_name_d["n"]: [n_S]},
                        {domain_to_name_d["cds"]: [cds_S]},
                        {domain_to_name_d["rflap"]: [rflap_S]},
                        {domain_to_name_d["term"]: [term_S]}
                    ]
                    
                    template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                    
                elif AGiloop == 6: # internal loop size
                    
                    ctRSD_part = hp5_S +\
                                 eo_S + r_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S2 + agl_S + ei_S + i_bm_S1 +\
                                 l_S + cl + rz_S +\
                                 s_S + i_th_S1 + i_bm_Sp1 + ei_Sp + i_th_S2 + i_bm_Sp2 + ei_Sp + o_th_S +\
                                 cp_S + n_S + cds_S + rflap_S +\
                                 term_S

                    rna_features = [
                        {domain_to_name_d["hp5"]: [hp5_S]},
                        {domain_to_name_d["e"]: [eo_S]},
                        {domain_to_name_d["r"]: [r_S]},
                        {domain_to_name_d["o_bm"]: [o_bm_S]},
                        {domain_to_name_d["c"]: [c_S]},
                        {domain_to_name_d["th"]: [o_th_Sp]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S2]},
                        {domain_to_name_d["agl"]: [agl_S]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S1]},
                        {domain_to_name_d["l"]: [l_S]},
                        {domain_to_name_d["cl"]: [cl]},
                        {domain_to_name_d["rz"]: [rz_S]},
                        {domain_to_name_d["spacer"]: [s_S]},
                        {domain_to_name_d["th"]: [i_th_S1]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp1]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [i_th_S2]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp2]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [o_th_S]},
                        {domain_to_name_d["cp"]: [cp_S]},
                        {domain_to_name_d["n"]: [n_S]},
                        {domain_to_name_d["cds"]: [cds_S]},
                        {domain_to_name_d["rflap"]: [rflap_S]},
                        {domain_to_name_d["term"]: [term_S]}
                    ]
                    
                    template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                        
            elif invert == 1:
                 
                if AGiloop == 5: # internal loop size
                    
                    ctRSD_part = hp5_S +\
                                 s_S + i_th_S1 + i_bm_Sp1 + ei_Sp + i_th_S2 + i_bm_Sp2 + ei_Sp + o_th_S +\
                                 cp_S + n_S + cds_S + rflap_S +\
                                 l_S + cl + rz_S + invL +\
                                 eo_S + r_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S2 + agl_S + i_th_Sp2[-1] + ei_S + i_bm_S1 +\
                                 term_S

                    rna_features = [
                        {domain_to_name_d["hp5"]: [hp5_S]},
                        {domain_to_name_d["s"]: [s_S]},
                        {domain_to_name_d["th"]: [i_th_S1]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp1]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [i_th_S2]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp2]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [o_th_S]},
                        {domain_to_name_d["cp"]: [cp_S]},
                        {domain_to_name_d["n"]: [n_S]},
                        {domain_to_name_d["cds"]: [cds_S]},
                        {domain_to_name_d["rflap"]: [rflap_S]},
                        {domain_to_name_d["l"]: [l_S]},
                        {domain_to_name_d["cl"]: [cl]},
                        {domain_to_name_d["rz"]: [rz_S]},
                        {domain_to_name_d["l"]: [invL]},
                        {domain_to_name_d["e"]: [eo_S]},
                        {domain_to_name_d["r"]: [r_S]},
                        {domain_to_name_d["o_bm"]: [o_bm_S]},
                        {domain_to_name_d["c"]: [c_S]},
                        {domain_to_name_d["th"]: [o_th_Sp]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S2]},
                        {domain_to_name_d["agl"]: [agl_S]},
                        {domain_to_name_d["th"]: [i_th_Sp2[-1]]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S1]},
                        {domain_to_name_d["term"]: [term_S]}
                    ]
                    
                    template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                        
                elif AGiloop == 6: # internal loop size
                    
                    ctRSD_part = hp5_S +\
                                 s_S + i_th_S1 + i_bm_Sp1 + ei_Sp + i_th_S2 + i_bm_Sp2 + ei_Sp + o_th_S +\
                                 cp_S + n_S + cds_S + rflap_S +\
                                 l_S + cl + rz_S + invL +\
                                 eo_S + r_S + o_bm_S + c_S + o_th_Sp + ei_S + i_bm_S2 + agl_S + ei_S + i_bm_S1 +\
                                 term_S

                    rna_features = [
                        {domain_to_name_d["hp5"]: [hp5_S]},
                        {domain_to_name_d["s"]: [s_S]},
                        {domain_to_name_d["th"]: [i_th_S1]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp1]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [i_th_S2]},
                        {domain_to_name_d["o_bm"]: [i_bm_Sp2]},
                        {domain_to_name_d["e"]: [ei_Sp]},
                        {domain_to_name_d["th"]: [o_th_S]},
                        {domain_to_name_d["cp"]: [cp_S]},
                        {domain_to_name_d["n"]: [n_S]},
                        {domain_to_name_d["cds"]: [cds_S]},
                        {domain_to_name_d["rflap"]: [rflap_S]},
                        {domain_to_name_d["l"]: [l_S]},
                        {domain_to_name_d["cl"]: [cl]},
                        {domain_to_name_d["rz"]: [rz_S]},
                        {domain_to_name_d["l"]: [invL]},
                        {domain_to_name_d["e"]: [eo_S]},
                        {domain_to_name_d["r"]: [r_S]},
                        {domain_to_name_d["o_bm"]: [o_bm_S]},
                        {domain_to_name_d["c"]: [c_S]},
                        {domain_to_name_d["th"]: [o_th_Sp]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S2]},
                        {domain_to_name_d["agl"]: [agl_S]},
                        {domain_to_name_d["e"]: [ei_S]},
                        {domain_to_name_d["i_bm"]: [i_bm_S1]},
                        {domain_to_name_d["term"]: [term_S]}
                    ]
                    template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
               
                
        elif CGs:
            '''
            ###################################################################
            COMPARATOR GATES
            ###################################################################
            '''
            
            dashcheckInp = re.compile(r"\-\d+\,")
            dashcheckInps = dashcheckInp.search(name.lower())
            
            dashcheckOut = re.compile(r"\-\d+\}")
            dashcheckOuts = dashcheckOut.search(name.lower())
            
            Indf = re.findall(r"[a-z]\d+",name.lower())
            dashIndf = re.findall(r"\-\d+",name.lower())
            
            if dashcheckInps and dashcheckOuts:
                i_th = Indf[0]
                o_th = Indf[1]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+dashIndf[1][1:]
                
            elif dashcheckInps:
                i_th = Indf[0]
                o_th = Indf[1][0]
                i_bm = 'D'+dashIndf[0][1:]
                o_bm = 'D'+Indf[1][1:]
                
            elif dashcheckOuts:
                i_th = Indf[0][0]
                o_th = Indf[1]
                i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+dashIndf[0][1:]
            
            else:
                i_th = Indf[0][0]
                o_th = Indf[1][0]
                i_bm = 'D'+Indf[0][1:]
                o_bm = 'D'+Indf[1][1:]
                
            if len(s) == 0:
                s_S = ''
            else:
                s_S = s_d[s][0]
            
            if len(cp) == 0:
                cp_S = ''
            else:
                cp_S = cp_d[cp][0]
                
            if len(n) == 0:
                n_S = ''
            else:
                n_S = n_d[n][0]
            
            
            if c == 0:
                c_S = ''
            else:
                if len(cp) == 0:
                    print('c domain included without complementary cp domain defined')
                    return('Function terminated, see error.')
                
                elif c > len(cp_S):
                    print('Length of specified c domain exceeds length of the defined complementary cp domain')
                    return('Function terminated, see error.')
                    
                else:
                    c_S = rc_seq(cp_S[0:c])[1:-1]
                    
            
            if len(CDS) == 0:
                cds_S = ''
            else:
                cds_S = cds_d[CDS][0]
                
            if len(rflap) == 0:
                rflap_S = ''
            else:
                rflap_S = rflap_d[rflap][0]
                
            if len(us) == 0:
                us_S = ''
            else:
                us_S = ''
                for n in range(len(us)):
                    us_S = us_S + us_d[us[m]][0]
                    
            if len(ds) == 0:
                ds_S = ''
            else:
                ds_S = ''
                for n in range(len(ds)):
                    ds_S = ds_S + ds_d[ds[m]][0]
                            
            # Extracted sequence domains
            I1_bm_S = O_BM_d[i_bm][0]
            I1_bm_Sp = rc_seq(O_BM_d[i_bm][0])[1:-1]
            I2_bm_S = O_BM_d[o_bm][0]
            I2_bm_Sp = rc_seq(O_BM_d[o_bm][0])[1:-1]
            i1_th_S = th_d[i_th][0]
            i2_th_S = th_d[o_th][0]
            cL_S = 'CGC'
            cL_Sp = 'GTG'
            # from optional inputs
            hp5_S = hp5_d[hp5][0]
            l_S = L_d[L][0]
            rz_S = Rz_d[Rz][0]
            term_S = term_d[term][0]
            prom_S = prom_d[prom][0]
            
            #Adding an indication of the Rz cleavage site for RNAs
            if rna == 1 and Rz[0].lower() != 'x':
                cl = '|'
            else:
                cl =''
                
            if invert == 1:
                print('To invert orientation, switch first and second input domains')
            
            ###################################################################
            # Stitching comparator gate sequences
            ###################################################################
            
            ctRSD_part = hp5_S + \
                         s_S + i1_th_S + c_S + I1_bm_Sp + cL_S + I2_bm_S + \
                         l_S + cl + rz_S + \
                         s_S + i2_th_S + I2_bm_Sp + cL_Sp + I1_bm_S + \
                         cp_S + n_S + cds_S + rflap_S +\
                         term_S

            rna_features = [
                {domain_to_name_d["hp5"]: [hp5_S]},
                {domain_to_name_d["s"]: [s_S]},
                {domain_to_name_d["th"]: [i1_th_S]},
                {domain_to_name_d["c"]: [c_S]},
                {domain_to_name_d["o_bm"]: [I1_bm_Sp]},
                {domain_to_name_d["cL"]: [cL_S]},
                {domain_to_name_d["i_bm"]: [I2_bm_S]},
                {domain_to_name_d["l"]: [l_S]},
                {domain_to_name_d["cl"]: [cl]},
                {domain_to_name_d["rz"]: [rz_S]},
                {domain_to_name_d["s"]: [s_S]},
                {domain_to_name_d["th"]: [i2_th_S]},
                {domain_to_name_d["o_bm"]: [I2_bm_Sp]},
                {domain_to_name_d["cL"]: [cL_Sp]},
                {domain_to_name_d["i_bm"]: [I1_bm_S]},
                {domain_to_name_d["cp"]: [cp_S]},
                {domain_to_name_d["n"]: [n_S]},
                {domain_to_name_d["cds"]: [cds_S]},
                {domain_to_name_d["rflap"]: [rflap_S]},
                {domain_to_name_d["term"]: [term_S]}
            ]
            
            template = create_template(ctRSD_part,prom_S,rna,us_S,ds_S,temp_len)
                           
        else:
            print('Incorrect Nomenclature')

        dna_features = [{"us": [us_S, str(us)]}] if us else []
        dna_features += [{domain_to_name_d["prom"]: [prom_S, str(prom)]}]
        dna_features += rna_features 
        dna_features += [{"ds": [ds_S, str(ds)]}] if ds else []

        rna_annotations = build_annotations(rna_features, domain_to_seq_to_partID_d)
        dna_annotations = build_annotations(dna_features, domain_to_seq_to_partID_d)

        genbank_dna = None
        genbank_rna = None
        if get_genbank:
            genbank_dna = write_genbank_text(
                template[0].replace("|", ""),  
                name.replace("{", "_").replace("}", "_").replace(",", "_"), 
                dna_annotations,
                rna=False
            )

            genbank_rna = write_genbank_text(
                template[1].replace("|", ""), 
                name.replace("{", "_").replace("}", "_").replace(",", "_"), 
                rna_annotations,
                rna=True
            )

        # Extract part IDs and types from annotations
        dna_part_IDs = [part['qualifiers']['note'] for part in dna_annotations]
        rna_part_IDs = [part['qualifiers']['note'] for part in rna_annotations]
        dna_part_types = [part['part_type'] for part in dna_annotations]
        rna_part_types = [part['part_type'] for part in rna_annotations]

        # Determine the orientation of each part based on the presence of the "_rc" suffix
        dna_part_orientations = ["reverse_complement" if part.endswith("_rc") else "forward" for part in dna_part_IDs]
        rna_part_orientations = ["reverse_complement" if part.endswith("_rc") else "forward" for part in rna_part_IDs]

        # Format part IDs: Remove the "_rc" suffix from part IDs
        dna_part_IDs = [part.replace("_rc", "") for part in dna_part_IDs]
        rna_part_IDs = [part.replace("_rc", "") for part in rna_part_IDs]
        
        return(
            template[0], 
            template[1], 
            dna_part_IDs,
            rna_part_IDs,
            dna_part_types,
            rna_part_types,
            dna_part_orientations,
            rna_part_orientations,
            genbank_dna,
            genbank_rna
        )


#############
#
# Compile Workbook
#
#############
EXPECTED_HEADERS = ["name", "Rz", "L", "term", "hp5", "prom", "eI", "eO", "s",
                    "invert", "invL", "agL", "AGiloop", "otype", "rna", "us", "ds",
                    "temp_len", "cp", "n", "c", "d", "CDS", "rflap"]


def process_workbook(file):
    # Load the workbook
    workbook = load_workbook(file, read_only=True, data_only=True)  #workbook = load_workbook(file.file)
    
    # Process sheets
    processed_sheets = [process_sheet(sheet) for sheet in workbook.worksheets]

    # Filter out None values (sheets that didn't match the expected headers)
    data = [sheet for sheet in processed_sheets if sheet is not None]

    # If no valid sheets were found, raise an error
    if not data:
        raise ValueError("No valid sheets found in the workbook. Ensure that the sheets have the expected headers: " + ", ".join(EXPECTED_HEADERS))

    # Make sure all sheets have the same length of data
    lengths = [len(sheet) for sheet in processed_sheets if sheet is not None]
    if len(set(lengths)) > 1:
        raise ValueError("Sheets in the workbook have different lengths of data. Ensure that all sheets have the same number of rows.")

    return data, lengths[0]

def process_sheet(sheet: Workbook):
    # Check if the sheet has the expected headers
    header = [cell.value for cell in sheet[1]]
    if header != EXPECTED_HEADERS:
        return None
    
    # Iterate through the rows in the sheet
    data = [dict(zip(EXPECTED_HEADERS, row)) for row in sheet.iter_rows(min_row=2, values_only=True)]

    return data

def compile_workbook_data(data, num_rows, domains_file):
    model = RSD_sim()
    model.load_domains(domains_file)

    name = []
    dna_template = []
    rna_template = []
    dna_part_IDs = []
    rna_part_IDs = []
    dna_part_types = []
    rna_part_types = []
    dna_part_orientations = []
    rna_part_orientations = []
    genbank_dna = []
    genbank_rna = []

    unique_parts = {}  # To track unique part IDs across all sheets

    for row_idx in range(num_rows):
        name_row = []
        dna_template_row = []
        rna_template_row = []
        dna_part_IDs_row = []
        rna_part_IDs_row = []
        dna_part_types_row = []
        rna_part_types_row = []
        dna_part_orientations_row = []
        rna_part_orientations_row = []
        genbank_dna_row = []
        genbank_rna_row = []

        for sheet_idx in range(len(data)):
            row = data[sheet_idx][row_idx]

            if row.get("name") is None or row.get("name").lower() == "none":
                continue  # Skip rows without a name

            result = model.ctRSD_seq_compile(
                name=row["name"],
                Rz=row["Rz"] or "Ro",
                L=row["L"] or "L",
                term=row["term"] or "T7t",
                hp5=row["hp5"] or "5hp",
                prom=row["prom"] or "P70a",
                eI=row["eI"] or '',
                eO=row["eO"] or '',
                s=row["s"] or '',
                invert=row["invert"] or 0,
                invL=row["invL"] or 'A',
                agL=row["agL"] or 'TA',
                AGiloop=row["AGiloop"] or 5,
                otype=row["otype"] or 1,
                rna=row["rna"] or 0,
                us=[x.strip()[1:-1] for x in row["us"][1:-1].split(',')] if row["us"] else [],
                ds=[x.strip()[1:-1] for x in row["ds"][1:-1].split(',')] if row["ds"] else [],
                temp_len=row["temp_len"] or 0,
                cp=row["cp"] or '',
                n=row["n"] or '',
                c=row["c"] or 0,
                d=row["d"] or '',
                CDS=row["CDS"] or '',
                rflap=row["rflap"] or '',
                get_genbank=1
            )

            if result == 'Function terminated, see error.':
                raise ValueError(f"Compilation failed for row {row_idx + 1} in sheet {sheet_idx + 1}. Please check the input parameters.")
            
            name_row.append(row["name"])
            dna_template_row.append(result[0])
            rna_template_row.append(result[1])
            dna_part_IDs_row.append(result[2])
            rna_part_IDs_row.append(result[3])
            dna_part_types_row.append(result[4])
            rna_part_types_row.append(result[5])
            dna_part_orientations_row.append(result[6])
            rna_part_orientations_row.append(result[7])
            genbank_dna_row.append(result[8])
            genbank_rna_row.append(result[9])

            for i in range(len(result[3])):
                part_id = result[3][i]
                part_type = result[5][i]
                part_sequence = model.domain_to_partID_to_seq_d.get(part_type, {}).get(part_id, '')
                unique_parts[part_id] = [part_type, part_sequence]  # Add RNA parts to the set of unique parts

        name.append(name_row)
        dna_template.append(dna_template_row)
        rna_template.append(rna_template_row)
        dna_part_IDs.append(dna_part_IDs_row)
        rna_part_IDs.append(rna_part_IDs_row)
        dna_part_types.append(dna_part_types_row)
        rna_part_types.append(rna_part_types_row)
        dna_part_orientations.append(dna_part_orientations_row)
        rna_part_orientations.append(rna_part_orientations_row)
        genbank_dna.append(genbank_dna_row)
        genbank_rna.append(genbank_rna_row)
    
    return {
        "name": name,
        "dna_template": dna_template,
        "rna_template": rna_template,
        "dna_part_IDs": dna_part_IDs,
        "rna_part_IDs": rna_part_IDs,
        "dna_part_types": dna_part_types,
        "rna_part_types": rna_part_types,
        "dna_part_orientations": dna_part_orientations,
        "rna_part_orientations": rna_part_orientations,
        "genbank_dna": genbank_dna,
        "genbank_rna": genbank_rna,
        "unique_parts": [{"part_id": k, "part_type": v[0], "part_sequence": v[1]} for k, v in unique_parts.items()]
    }
