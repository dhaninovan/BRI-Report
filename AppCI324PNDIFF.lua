-- Diff Report : Find and calculate the difference between files CI324 PN FDS Monthly Trial Balance (CSV Format from Portal DWH)
-- Interpreter : (Novan's Modified)lua.exe
-- 10:03 PM Sunday, February 10, 2019 Jakarta City

-- Format output report
-- 1. 			No Rekening c6
-- 2. 			Tipe Tenor c5 c14
-- 3. 			Jatuh Tempo c12
-- 4. 			Nama c9
-- 5. 			Saldo Awal c8
-- 6. 			Saldo Akhir
-- 7. 			Delta
-- 8. 			Officer c17 c18
			
list_acc = {}
OUTPUT_FILE = "CI324PNDIFF"

function format_account(s)
    local rek_len = string.len(s)
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

res, ReportFileName1, ReportFileName2, output_sep, limit_res = iup.GetParam("Pilih Report CI324 PN dalam Format CSV (Sumber: DWH)", nil, [=[
Sumber Data: %m\n
Report Posisi Awal: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Report Posisi Akhir: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Output Separator: %l|,|;|\n
Limit Result: %l|10|20|30|50|100|\n
]=]
,"1. Buka Aplikasi BRISIM (https://brisim.bri.co.id)\n2. Pilih: DWH Reports\n3. Pilih: Critical Report\n4. Pilih: Table\n5. Pilih CI324(PN) - FDS MONTHLY TRIAL BALANCE BY PRODUCT TYPE(1 ROW)\n6. Download dan Save dalam format CSV", "C:\\Lua\\data\\20201231 CI324Modif.csv","C:\\Lua\\data\\20210131 CI324Modif.csv",0,1)

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
			acc_type_tenor = f[5].." - "..f[14]
			acc_maturity = f[12]
			acc_name = f[9]
			acc_officer = f[16].." - "..f[17]
			acc_balance = string.gsub(string.sub(f[8], 1, #f[8]-3), ",", "")
			acc_balance = tonumber(acc_balance)
			list_acc[acc_no] = {acc_type_tenor, acc_maturity, acc_name, acc_balance, 0, -acc_balance, acc_officer}
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
			acc_type_tenor = f[5].." - "..f[14]
			acc_maturity = f[12]
			acc_name = f[9]
			acc_officer = f[16].." - "..f[17]
			acc_balance = string.gsub(string.sub(f[8], 1, #f[8]-3), ",", "")
			acc_balance = tonumber(acc_balance)
			if list_acc[acc_no] then
				list_acc[acc_no][5] = acc_balance
				list_acc[acc_no][6] = list_acc[acc_no][5] - list_acc[acc_no][4]
			else
				list_acc[acc_no] = {acc_type_tenor, acc_maturity, acc_name, 0, acc_balance, acc_balance, acc_officer}
			end
		end
	end
	no = no + 1
end

print('Sorting descending')
sorted_list_acc = {}
for k, v in pairs(list_acc) do
	table.insert(sorted_list_acc, {k, v[1], v[2], v[3], v[4], v[5], v[6], v[7]}) 
end
table.sort(sorted_list_acc, function(a,b) return a[7]<b[7] end)
print("Processing "..tostring(#sorted_list_acc).." rekening")

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
fo:write('No Rek'..output_sep..'CIF'..output_sep..'Nama'..output_sep..'Posisi Awal'..output_sep..'Posisi Akhir'..output_sep..'Delta'..output_sep..'Officer\n')
fo2:write([=[
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 
<title>BRI Reporting Tool: Delta Deposito CI324</title>
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
 <h2>&nbsp;&nbsp;&nbsp;DEPOSITO YANG CAIR (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>Tenor</th>
			<th>Jatuh Tempo</th>
			<th>Nama</th>
			<th>Saldo Awal<br>]=]..posisi_report1..[=[</th>
			<th>Saldo Akhir<br>]=]..posisi_report2..[=[</th>
			<th>Delta</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(sorted_list_acc) do	
	if ii < limit_res then
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td>"..v[4].."</td><td align='right'>"..format_number(v[5]).."</td><td align='right'>"..format_number(v[6]).."</td><td class='minus' align='right'>"..format_number(v[7]).."</td><td align='right'>"..v[8].."</td><td/></tr>\n")
	elseif ii == limit_res then
		fo2:write([=[
	</tbody>
 </table>
 <hr></br><h2>&nbsp;&nbsp;&nbsp;DEPOSITO YANG BARU (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>Tenor</th>
			<th>Jatuh Tempo</th>
			<th>Nama</th>
			<th>Saldo Awal<br>]=]..posisi_report1..[=[</th>
			<th>Saldo Akhir<br>]=]..posisi_report2..[=[</th>
			<th>Delta</th>
			<th>Officer</th>
			<th width=200>Keterangan</th>
		</tr>
	</thead>
	<tbody>
		]=])
	end
	
	if (#sorted_list_acc - ii) <= limit_res then
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td>"..v[4].."</td><td align='right'>"..format_number(v[5]).."</td><td align='right'>"..format_number(v[6]).."</td><td class='plus' align='right'>"..format_number(v[7]).."</td><td align='right'>"..v[8].."</td><td/></tr>\n")
	end
	
	fo:write(v[1]..output_sep..v[2]..output_sep..'"'..v[3]..'"'..output_sep..v[4]..output_sep..v[5]..output_sep..v[6]..output_sep..v[7]..output_sep..v[8]..'\n')
	ii = ii + 1
end
fo2:write([=[
	</tbody>
 </table>
 &nbsp;&nbsp;&nbsp;&nbsp;Data selengkapnya: <a href=']=]..OUTPUT_FILE..[=[.csv'>]=]..OUTPUT_FILE..[=[.csv</a>
 <hr>
 <div align='center' style='font-size:smaller'>BRI Reporting Tool: CI324 Diff<br>Copyright &copy;2013, <b>Dhani Novan</b> (dhani_novan@bri.co.id)</div><br/>
</body>
</html>
]=])
fo2:close()
fo:close()
print('=== Done in '..(os.clock()-t1)..' ===')
os.execute(OUTPUT_FILE..".htm")
--os.execute("pause")


