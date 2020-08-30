-- Diff Report : Find and calculate the difference between files LW321 Loan (CSV Format from Portal DWH)
-- Interpreter : (Novan's Modified)lua.exe
-- 5:46 AM Sunday, February 24, 2019 Jakarta City

-- LNTYPE : RITKOM : DA,DH,DL,EB,NM,T6,UZ,VZ
-- LNTYPE : PROGRAM : RW,RV,RK,PC
-- LNTYPE : KONSUMER : WP,WM,ZU,PH,HT,FH,K3,K4


--Data sumber yang dibutuhkan dari LW321
-- [No Kolom] [Keterangan] [variable lua]
--5 LN Type : acc_type
--6 Nomor rekening : acc_no
--7 Nama Debitur : acc_name
--8 Plafond : acc_plafond
--13 Tgl Realisasi : acc_realisasi_dt
--14 Tgl Jatuh tempo : acc_jt_dt
--15 Jangka Waktu : acc_jw
--16 Flag Restruk : acc_flag_restruk
--18 19 20 21 22 Kolektibilitas : acc_kolek
--18 19 20 21 22 Outstanding : acc_os
--26 PN : acc_officer
--27 Nama PN : acc_officer

--Raw Output (key = Nomor rekening):
--  1. LN Type acc_type
--  2. Plafond acc_plafond
--  3. Tgl Realisasi acc_realisasi_dt
--  4. Tgl Jatuh tempo acc_jt_dt
--  5. Jangka Waktu acc_jw
--  6. Flag Restruk acc_flag_restruk
--  7. Nama Debitur acc_name
--  8. Outstanding Awal *acc_os
--  9. Outstanding Akhir *acc_os
-- 10. Delta Outstanding *acc_os
-- 11. Kolek Awal acc_kolek
-- 12. Kolek Akhir acc_kolek
-- 13. PN + Nama acc_officer

--Delta Outstanding Report:
--1 Nomor rekening [key]
--2 LN Type  fo[1]
--3 Plafond  fo[2]
--4 Tgl Realisasi	fo[3]
--5 Nama Debitur fo[7] 
--6 Outstanding Awal fo[8] 
--7 Outstanding Akhir fo[9]
--8 Delta fo[10]
--9 PN + Nama fo[13]

--Pemburukan DPK/NPL Baru:
--1 Nomer rekening [key]
--2 LN Type fo[1]
--3 Tgl Realisasi fo[3]
--4 Tgl Jatuh Tempo fo[4]
--5 Restruk fo[6]
--6 Nama fo[7]
--7 Outstanding fo[9]
--8 Keterangan Kolek fo[11] fo[12]
--9 Officer fo[13]

list_acc = {}
OUTPUT_FILE = "LW321DIFF"

function format_account(s)
    local rek_len = string.len(s)
    if rek_len > 15 then
		s = string.sub(s, rek_len-15+1)
		rek_len = 15
    end
    if rek_len >= 12 and rek_len <=15 then
        s = string.rep("0",15-rek_len)..s
        s = string.sub(s,1,-12).."-"..string.sub(s,-11,-10).."-"..string.sub(s,-9,-4).."-"..string.sub(s,-3,-2).."-"..string.sub(s,-1,-1)
    end
    return s
end

function format_number(v)
	local s
	local unary
	
	if v < 0 then
		--s = string.format("%d", math.floor(-v))
		s = tostring(-v)
		unary = "-"
	else
		--s = string.format("%d", math.floor(v))
		s = tostring(v)
		unary = ""
	end
    
    local pos = string.len(s) % 3

    if pos == 0 then pos = 3 end
    return unary..string.sub(s, 1, pos).. string.gsub(string.sub(s, pos+1), "(...)", ".%1")
--	return v
end

function FindFirstSeparator(line)
	local sep,c
	
	sep = ','
	for c in line:gmatch(".") do
		if c == ',' then
			sep = ',' break
		elseif c == ';' then
			sep = ';' break
		elseif c == ':' then
			sep = ':' break
		elseif c == '|' then
			sep = '|' break
		end
	end
	return sep
end

res, ReportFileName1, ReportFileName2, output_sep, limit_res = iup.GetParam("Pilih Report LW321 dalam Format CSV (Sumber: DWH)", nil, [=[
Report Posisi Awal: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Report Posisi Akhir: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Output Separator: %l|,|;|\n
Limit Result: %l|10|20|30|50|100|\n
]=]
,"","",0,1)

if ReportFileName1 == "" or ReportFileName2 == "" then
	print("Please select two reports to be compared")
	os.execute("pause")
	os.exit(-1)
end

-- convert Unicode to ANSI
print('Converting '..ReportFileName1..' to ANSI encoding')
os.execute('type "'..ReportFileName1..'" > '..'tmp.csv')
os.remove(ReportFileName1)
os.rename('tmp.csv', ReportFileName1)

print('Converting '..ReportFileName2..' to ANSI encoding')
os.execute('type "'..ReportFileName2..'" > '..'tmp.csv')
os.remove(ReportFileName2)
os.rename('tmp.csv', ReportFileName2)

-- Load first data into table list_acc
t1 = os.clock()
print('Loading data from '..ReportFileName1)
no = 1
sep = ','
posisi_report1 = ''
for line in io.lines(ReportFileName1) do
	-- process header
	if no == 1 then
		sep = FindFirstSeparator(line)
	else
		-- process data
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			posisi_report1 = f[1]
			acc_no = f[6]
			acc_type = f[5]
			acc_name = f[7]
			acc_plafond = f[8]
			acc_realisasi_dt = f[13]
			acc_jt_dt = f[14]
			acc_jw = f[15]
			acc_flag_restruk = f[16]
			acc_officer = f[26].." - "..f[27]
			if f[18] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[18], 1, #f[18]-3), ",", "")
				acc_kolek = 1
			elseif f[19] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[19], 1, #f[19]-3), ",", "")
				acc_kolek = 2
			elseif f[20] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[20], 1, #f[20]-3), ",", "")
				acc_kolek = 3
			elseif f[21] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[21], 1, #f[21]-3), ",", "")
				acc_kolek = 4
			elseif f[22] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[22], 1, #f[22]-3), ",", "")
				acc_kolek = 5
			end			
			acc_os = tonumber(acc_os)
			list_acc[acc_no] = {acc_type, acc_plafond, acc_realisasi_dt, acc_jt_dt, acc_jw, acc_flag_restruk, acc_name, acc_os, 0, -acc_os, acc_kolek, acc_kolek, acc_officer}
		end
	end
	no = no + 1
end

-- Update table list_acc with lastest balance from second data
print('Loading data from '..ReportFileName2)
no = 1
sep = ','
posisi_report2 = ''
for line in io.lines(ReportFileName2) do
	-- only process line begin with number, skipping header
	if no == 1 then
		sep = FindFirstSeparator(line)
	else
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			posisi_report2 = f[1]
			acc_no = f[6]
			acc_type = f[5]
			acc_name = f[7]
			acc_plafond = f[8]
			acc_realisasi_dt = f[13]
			acc_jt_dt = f[14]
			acc_jw = f[15]
			acc_flag_restruk = f[16]
			acc_officer = f[26].." - "..f[27]
			if f[18] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[18], 1, #f[18]-3), ",", "")
				acc_kolek = 1
			elseif f[19] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[19], 1, #f[19]-3), ",", "")
				acc_kolek = 2
			elseif f[20] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[20], 1, #f[20]-3), ",", "")
				acc_kolek = 3
			elseif f[21] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[21], 1, #f[21]-3), ",", "")
				acc_kolek = 4
			elseif f[22] ~= "0.00" then
				acc_os = string.gsub(string.sub(f[22], 1, #f[22]-3), ",", "")
				acc_kolek = 5
			end			
			acc_os = tonumber(acc_os)
			if list_acc[acc_no] then
				list_acc[acc_no][9] = acc_os
				list_acc[acc_no][10] = list_acc[acc_no][9] - list_acc[acc_no][8]
				list_acc[acc_no][12] = acc_kolek
				list_acc[acc_no][2] = acc_plafond
			else
				list_acc[acc_no] = {acc_type, acc_plafond, acc_realisasi_dt, acc_jt_dt, acc_jw, acc_flag_restruk, acc_name, 0, acc_os, acc_os, acc_kolek, acc_kolek, acc_officer}
			end
		end
	end
	no = no + 1
end

print('Sorting descending by delta OS')
sorted_list_acc = {}
for k, v in pairs(list_acc) do
	-- table.insert(sorted_list_acc, {k, v[1], v[2], v[3], v[7], v[8], v[9], v[10], v[13]}) 
	table.insert(sorted_list_acc, {k, v[1].." - "..v[11]..v[12], v[2], v[3], v[7], v[8], v[9], v[10], v[13]}) 
end
table.sort(sorted_list_acc, function(a,b) return a[8]<b[8] end)
print("Processing "..tostring(#sorted_list_acc).." accounts")

print('Sorting descending by DPK/NPL')
sorted_list_acc_newNPL = {}	-- new NPL
sorted_list_acc_newDPK = {}	-- new DPK
sorted_list_acc_NPL = {}	-- existing NPL
sorted_list_acc_DPK = {}	-- existing DPK
for k, v in pairs(list_acc) do
	if v[11] < 3 and v[12] >= 3 then	-- dari kolek 1,2 pindah ke kolek 3,4,5
		table.insert(sorted_list_acc_newNPL, {k, v[1], v[3], v[4], v[6], v[7], v[9], v[11].." ke "..v[12], v[13]}) 
	elseif v[11] == 1 and v[12] == 2 then -- dari kolek 1 pindah ke kolek 2
		table.insert(sorted_list_acc_newDPK, {k, v[1], v[3], v[4], v[6], v[7], v[9], v[11].." ke "..v[12], v[13]})
	elseif v[12] == 2 then
		table.insert(sorted_list_acc_DPK, {k, v[1], v[3], v[4], v[6], v[7], v[9], v[12], v[13]})
	elseif v[12] >= 3 then
		table.insert(sorted_list_acc_NPL, {k, v[1], v[3], v[4], v[6], v[7], v[9], v[12], v[13]})
	end
end
table.sort(sorted_list_acc_newNPL, function(a,b) return a[7]>b[7] end)
print("Processing New NPL "..tostring(#sorted_list_acc_newNPL).." accounts")
table.sort(sorted_list_acc_newDPK, function(a,b) return a[7]>b[7] end)
print("Processing New DPK "..tostring(#sorted_list_acc_newDPK).." accounts")
table.sort(sorted_list_acc_NPL, function(a,b) return a[7]>b[7] end)
print("Processing NPL "..tostring(#sorted_list_acc_NPL).." accounts")
table.sort(sorted_list_acc_DPK, function(a,b) return a[7]>b[7] end)
print("Processing DPK "..tostring(#sorted_list_acc_DPK).." accounts")

print('Writing '..OUTPUT_FILE)
if output_sep == 0 then
	output_sep = ','
elseif output_sep == 1 then
	output_sep = ';'
end
if limit_res == 0 then
	limit_res = 10
elseif limit_res == 1 then
	limit_res = 20
elseif limit_res == 2 then
	limit_res = 30
elseif limit_res == 3 then
	limit_res = 50
elseif limit_res == 4 then
	limit_res = 100
end
ii = 0
fo = io.open(OUTPUT_FILE..".csv", "w")
fo2 = io.open(OUTPUT_FILE..".htm", "w")
fo:write('No Rek'..output_sep..'Type'..output_sep..'Plafond'..output_sep..'Realisasi'..output_sep..'Nama'..output_sep..'OS Awal'..output_sep..'OS Akhir'..output_sep..'Delta'..output_sep..'Officer\n')
fo2:write([=[
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 
<title>BRI Reporting Tool: Delta Tabungan LW321</title>
<style type="text/css"> 
body, html  { height: 100%; }
html, body, div, span, applet, object, iframe,
/*h1, h2, h3, h4, h5, h6,*/ p, blockquote, pre,
a, abbr, acronym, address, big, cite, code,
del, dfn, em, font, img, ins, kbd, q, s, samp,
small, strike, strong, sub, sup, tt, var,
b, u, i, center,
dl, dt, dd, ol, ul, li,
fieldset, form, label, legend,
table, caption, tbody, tfoot, thead, tr, th, td {
	margin: 0;
	padding: 0;
	border: 0;
	outline: 0;
	font-size: 100%;
	vertical-align: baseline;
	background: transparent;
}
body { line-height: 1; }
ol, ul { list-style: none; }
blockquote, q { quotes: none; }
blockquote:before, blockquote:after, q:before, q:after { content: ''; content: none; }
:focus { outline: 0; }
del { text-decoration: line-through; }
table {border-spacing: 0; }
 
/*------------------------------------------------------------------ */
 body{
	font-family:Arial, Helvetica, sans-serif;
	margin:0 auto;
}
a:link {
	color: #666;
	font-weight: bold;
	text-decoration:none;
}
a:visited {
	color: #666;
	font-weight:bold;
	text-decoration:none;
}
a:active,
a:hover {
	color: #bd5a35;
	text-decoration:underline;
}
 
table a:link {
	color: #666;
	font-weight: bold;
	text-decoration:none;
}
table a:visited {
	color: #999999;
	font-weight:bold;
	text-decoration:none;
}
table a:active,
table a:hover {
	color: #bd5a35;
	text-decoration:underline;
}
table {
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
	margin:20px;
	border: 1px solid;
}
table th {
	padding:10px 25px 10px 25px; 
	background: #000000	;
	border-left: 1px solid #a0a0a0;
	color:#ffffff;
}
table th:first-child {
	border-left: 0;
}
table tr {
	text-align: center;
	padding-left:20px;
}
table td:first-child {
	padding-left:20px;
	border-left: 0;
}
table td {
	padding:10px;
	border-top: 0px solid #ffffff;
	border-bottom:1px solid #a0a0a0;
	border-left: 1px solid #a0a0a0;
	
	background: #ffffff;
}
table tr.even td {
	background: #eeeeee;
}
table tr:hover td {
	background: #bbbbff;
}
table td.minus {
	color: #ff0000;
}
table td.plus {
	color: #008800;
} 
</style>
 </head>
 <body>
 <h2>&nbsp;&nbsp;&nbsp;Pinjaman Lunas atau disetor(Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>Type</th>
			<th>Plafond</th>
			<th>Realisasi</th>
			<th>Nama</th>
			<th>OS Awal<br>]=]..posisi_report1..[=[</th>
			<th>OS Akhir<br>]=]..posisi_report2..[=[</th>
			<th>Delta</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc) do	
	if ii < limit_res then
		fo2:write("<tr><td>"
		..format_account(v[1]).."</td><td>"
		..v[2].."</td><td>"
		..v[3].."</td><td>"
		..v[4].."</td><td>"
		..v[5].."</td><td align='right'>"
		..format_number(v[6]).."</td><td align='right'>"
		..format_number(v[7]).."</td><td class='minus' align='right'>"
		..format_number(v[8]).."</td><td align='right'>"
		..v[9].."</td><td/></tr>\n")
	elseif ii == limit_res then
		fo2:write([=[
	</tbody>
 </table>
 <hr></br><h2>&nbsp;&nbsp;&nbsp;PINJAMAN BARU ATAU PENCAIRAN (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>Type</th>
			<th>Plafond</th>
			<th>Realisasi</th>
			<th>Nama</th>
			<th>OS Awal<br>]=]..posisi_report1..[=[</th>
			<th>OS Akhir<br>]=]..posisi_report2..[=[</th>
			<th>Delta</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
		]=])
	end
	
	if (#sorted_list_acc - ii) <= limit_res then
		fo2:write("<tr><td>"
		..format_account(v[1]).."</td><td>"
		..v[2].."</td><td>"
		..v[3].."</td><td>"
		..v[4].."</td><td>"
		..v[5].."</td><td align='right'>"
		..format_number(v[6]).."</td><td align='right'>"
		..format_number(v[7]).."</td><td class='plus' align='right'>"
		..format_number(v[8]).."</td><td align='right'>"
		..v[9].."</td><td/></tr>\n")
	end
	
	fo:write(v[1]..output_sep..v[2]..output_sep..'"'..v[3]..'"'..output_sep..v[4]..output_sep..v[5]..output_sep..v[6]..output_sep..v[7]..output_sep..v[8]..output_sep..v[9]..'\n')
	ii = ii + 1
end
fo2:write([=[
	</tbody>
 </table>
&nbsp;&nbsp;&nbsp;&nbsp;Data selengkapnya: <a href=']=]..OUTPUT_FILE..[=[.csv'>]=]..OUTPUT_FILE..[=[.csv</a>
 ]=])
 
 ------------------------ New NPL Report
 fo2:write([=[<hr></br><h2>&nbsp;&nbsp;&nbsp;NPL Baru (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>Rekening</th>
			<th>Type</th>
			<th>Realisasi</th>
			<th>Jatuh Tempo</th>
			<th>Restruk</th>
			<th>Nama</th>
			<th>Baki Debet</th>
			<th>Kolektiblitas</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc_newNPL) do
	fo2:write("<tr><td>"
	..format_account(v[1]).."</td><td>"
	..v[2].."</td><td>"
	..v[3].."</td><td>"
	..v[4].."</td><td>"
	..v[5].."</td><td>"
	..v[6].."</td><td class='minus' align='right'>"
	..format_number(v[7]).."</td><td>"
	..v[8].."</td><td>"
	..v[9].."</td><td/></tr>\n")	
end
fo2:write([=[
	</tbody>
 </table>
 ]=])
 
 ------------------------ New DPK Report
 fo2:write([=[<hr></br><h2>&nbsp;&nbsp;&nbsp;DPK Baru (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>Rekening</th>
			<th>Type</th>
			<th>Realisasi</th>
			<th>Jatuh Tempo</th>
			<th>Restruk</th>
			<th>Nama</th>
			<th>Baki Debet</th>
			<th>Kolektiblitas</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc_newDPK) do
	if v[2] ~= "FV" and v[2] ~= "WL" then
	fo2:write("<tr><td>"
	..format_account(v[1]).."</td><td>"
	..v[2].."</td><td>"
	..v[3].."</td><td>"
	..v[4].."</td><td>"
	..v[5].."</td><td>"
	..v[6].."</td><td class='minus' align='right'>"
	..format_number(v[7]).."</td><td>"
	..v[8].."</td><td>"
	..v[9].."</td><td/></tr>\n")
	end
end
fo2:write([=[
	</tbody>
 </table>
 ]=])
 
 ------------------------ NPL Report
 fo2:write([=[<hr></br><h2>&nbsp;&nbsp;&nbsp;Rincian Eksisting NPL (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>Rekening</th>
			<th>Type</th>
			<th>Realisasi</th>
			<th>Jatuh Tempo</th>
			<th>Restruk</th>
			<th>Nama</th>
			<th>Baki Debet</th>
			<th>Kolektiblitas</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc_NPL) do
	fo2:write("<tr><td>"
	..format_account(v[1]).."</td><td>"
	..v[2].."</td><td>"
	..v[3].."</td><td>"
	..v[4].."</td><td>"
	..v[5].."</td><td>"
	..v[6].."</td><td class='minus' align='right'>"
	..format_number(v[7]).."</td><td>"
	..v[8].."</td><td>"
	..v[9].."</td><td/></tr>\n")	
end
fo2:write([=[
	</tbody>
 </table>
 ]=])
  
 ------------------------ DPK Report
 fo2:write([=[<hr></br><h2>&nbsp;&nbsp;&nbsp;Rincian Eksisting DPK (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>Rekening</th>
			<th>Type</th>
			<th>Realisasi</th>
			<th>Jatuh Tempo</th>
			<th>Restruk</th>
			<th>Nama</th>
			<th>Baki Debet</th>
			<th>Kolektiblitas</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc_DPK) do
	fo2:write("<tr><td>"
	..format_account(v[1]).."</td><td>"
	..v[2].."</td><td>"
	..v[3].."</td><td>"
	..v[4].."</td><td>"
	..v[5].."</td><td>"
	..v[6].."</td><td class='minus' align='right'>"
	..format_number(v[7]).."</td><td>"
	..v[8].."</td><td>"
	..v[9].."</td><td/></tr>\n")	
end
fo2:write([=[
	</tbody>
 </table>
 ]=])
 
---------------------- End of Report
 fo2:write([=[
 <hr><div align='center' style='font-size:smaller'>BRI Reporting Tool: LW321 Diff<br>Copyright &copy;2013, <b>Dhani Novan</b> (dhani_novan@bri.co.id)</div><br/>
</body>
</html>
]=])
fo2:close()
fo:close()
print('=== Done in '..(os.clock()-t1)..' ===')
os.execute(OUTPUT_FILE..".htm")
--os.execute("pause")


