-- Diff Report : Find and calculate the difference between files DI321PN Current Account Balance (CSV Format from DWH BRISIM)
-- Interpreter : (Novan's Modified)lua.exe
-- 7:57 PM Thursday, April 25, 2013 Raha City
-- 8:45 PM Thursday, Mei 07, 2019 Rawamangun City

-- Format Column CSV
--  1. periode
--  2. uker code
--  3. curr code
--  4. curr desc
--  5. cifno
--  6. account number
--  7. status
--  8. short name
--  9. dlt
-- 10. balance
-- 11. avail balance
-- 12. limit
-- 13. cr int
-- 14. dr int
-- 15. commintment
-- 16. avrg balance
-- 17. PN PENGELOLA
-- 18. NAMA PENGELOLA

list_acc = {}
OUTPUT_FILE = "DI321PNDIFF"
f_lines = nil

-- false: hide internal account number -9x-
-- true: show internal account number -9x-
INCLUDE_IA = false

function format_account(s)
    local rek_len = string.len(s)
    if rek_len >= 12 and rek_len <=15 then
        s = string.rep("0",15-rek_len)..s
        s = string.sub(s,1,-12).."-"..string.sub(s,-11,-10).."-"..string.sub(s,-9,-4).."-"..string.sub(s,-3,-2).."-"..string.sub(s,-1,-1)
    end
    return s
end

function ReadRegistry(key, value)
	local fi, data, content, data_type
	
	fi = io.popen(string.format('reg QUERY "%s" /v %s', key, value))
	data = nil
	if fi then
		content = fi:read("*a")
		data_type, data = content:match(value..'%s+(%S+)%s+(.+)\n\n')
		fi:close()
	end

	return data_type, data
end

function format_number(v)
	local s
	local unary, sep_thousand
	
	if v < 0 then
		s = tostring(-v)
		unary = "-"
	else
		s = tostring(v)
		unary = ""
	end
    
    local pos = string.len(s) % 3
	data_type, thousand_sep = ReadRegistry('HKCU\\Control Panel\\International', 'sThousand')
	if thousand_sep == nil then thousand_sep = '.' end
    if pos == 0 then pos = 3 end
    return unary..string.sub(s, 1, pos).. string.gsub(string.sub(s, pos+1), "(...)", thousand_sep.."%1")
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

res, dummy, ReportFileName1, ReportFileName2, limit_res = iup.GetParam("Pilih Report DI321 PN dalam Format CSV (Sumber: DWH)", nil, [=[
Sumber Data: %m\n
Report Posisi Awal: %f[OPEN|*DI321*.csv;*DI321*.gz|CURRENT|NO|NO]\n
Report Posisi Akhir: %f[OPEN|*DI321*.csv;*DI321*.gz|CURRENT|NO|NO]\n
Limit Result: %l|10|20|30|50|100|\n
]=]
,"1. Buka Aplikasi BRISIM (https://brisim.bri.co.id)\n2. Pilih: DWH Reports\n3. Pilih: Critical Report\n4. Pilih: Table\n5. Pilih DI321(PN) - CURRENT ACCOUNT MONTHLY TRIAL BALANCE - ACTIVE(1 ROW)\n6. Download dan Save dalam format CSV", "C:\\Lua\\data\\20201231 DI321 PN PENGELOLAH.csv","C:\\Lua\\data\\20210130 DI321 PN PENGELOLAH.csv",1)

data_type, output_sep = ReadRegistry('HKCU\\Control Panel\\International', 'sList')
data_type, decimal_sep = ReadRegistry('HKCU\\Control Panel\\International', 'sDecimal')
if output_sep == decimal_sep then output_sep =';' end
if output_sep == nil then output_sep = ',' end

if ReportFileName1 == "" or ReportFileName2 == "" then
	print("Please select two reports to be compared")
	os.execute("pause")
	os.exit(-1)
end

-- convert Unicode to ANSI
if ReportFileName1:match('%.gz$') == nil then
print('Converting '..ReportFileName1..' to ANSI encoding')
os.execute('type "'..ReportFileName1..'" > '..'tmp.csv')
os.remove(ReportFileName1)
os.rename('tmp.csv', ReportFileName1)
f_lines = io.lines
else
f_lines = gzio.lines
end

if ReportFileName1:match('%.gz$') == nil then
print('Converting '..ReportFileName2..' to ANSI encoding')
os.execute('type "'..ReportFileName2..'" > '..'tmp.csv')
os.remove(ReportFileName2)
os.rename('tmp.csv', ReportFileName2)
end

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
			acc_cif = f[5]
			acc_name = f[8]
			acc_balance = string.gsub(string.sub(f[10], 1, #f[10]-3), ",", "")
			acc_balance = tonumber(acc_balance)
			acc_officer = f[17].." - "..f[18]
			if INCLUDE_IA or (string.sub(acc_no,-3,-3) ~= "9") then
				list_acc[acc_no] = {acc_cif, acc_name, acc_balance, 0, -acc_balance, acc_officer}
			end
		end
	end
	no = no + 1
end

-- Update table list_acc with lastest balance from second data
print('Loading data from '..ReportFileName2)
no = 1
sep = ','
posisi_report2 = ''
fo = io.open(OUTPUT_FILE.."_NEW.csv", "w")
fo:write('Rekening'..output_sep..'Nama'..output_sep..'Saldo'..output_sep..'PN_Pengelola\n')
for line in f_lines(ReportFileName2) do
	-- only process line begin with number, skipping header
	if no == 1 then
		sep = FindFirstSeparator(line)
	else
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			posisi_report2 = f[1]
			acc_no = f[6]
			acc_cif = f[5]
			acc_name = f[8]
			acc_balance = string.gsub(string.sub(f[10], 1, #f[10]-3), ",", "")
			acc_balance = tonumber(acc_balance)
			acc_officer = f[17].."-"..f[18]
			if INCLUDE_IA or (string.sub(acc_no,-3,-3) ~= "9") then
				if list_acc[acc_no] then
					list_acc[acc_no][4] = acc_balance
					list_acc[acc_no][5] = list_acc[acc_no][4] - list_acc[acc_no][3]
					list_acc[acc_no][6] = acc_officer
				else
					list_acc[acc_no] = {acc_cif, acc_name, acc_balance, 0, -acc_balance, acc_officer}
					fo:write(string.format('%s%s%s%s%s%s%s\n', 
						format_account(acc_no), output_sep,
						acc_name, output_sep,
						format_number(acc_balance), output_sep,
						acc_officer))
				end
			end
		end
	end
	no = no + 1
end
fo:close()

print('Sorting descending')
sorted_list_acc = {}
for k, v in pairs(list_acc) do
	table.insert(sorted_list_acc, {k, v[1], v[2], v[3], v[4], v[5], v[6]}) 
end
table.sort(sorted_list_acc, function(a,b) return a[6]<b[6] end)
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
 
<title>BRI Reporting Tool: Delta Giro DI321</title>
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
 <h2>&nbsp;&nbsp;&nbsp;REKENING YANG SALDONYA TURUN (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>CIF</th>
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
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td align='right'>"..format_number(v[4]).."</td><td align='right'>"..format_number(v[5]).."</td><td class='minus' align='right'>"..format_number(v[6]).."</td><td>"..v[7].."</td><td/></tr>\n")
	elseif ii == limit_res then
		fo2:write([=[
	</tbody>
 </table>
 <hr></br><h2>&nbsp;&nbsp;&nbsp;REKENING YANG SALDONYA NAIK (Periode ]=]..posisi_report1..[=[ - ]=]..posisi_report2..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>No Rekening</th>
			<th>CIF</th>
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
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td align='right'>"..format_number(v[4]).."</td><td align='right'>"..format_number(v[5]).."</td><td class='plus' align='right'>"..format_number(v[6]).."</td><td>"..v[7].."</td><td/></tr>\n")
	end
	
	fo:write(v[1]..output_sep..v[2]..output_sep..'"'..v[3]..'"'..output_sep..v[4]..output_sep..v[5]..output_sep..v[6]..output_sep..v[7]..'\n')
	ii = ii + 1
end
fo2:write([=[
	</tbody>
 </table>
 &nbsp;&nbsp;&nbsp;&nbsp;Data selengkapnya: <a href=']=]..OUTPUT_FILE..[=[.csv'>]=]..OUTPUT_FILE..[=[.csv</a>
 <hr>
 <div align='center' style='font-size:smaller'>BRI Reporting Tool: DI321 Diff<br>Copyright &copy;2013, <b>Dhani Novan</b> (dhani_novan@bri.co.id)</div><br/>
</body>
</html>
]=])
fo2:close()
fo:close()
print('=== Done in '..(os.clock()-t1)..' ===')
os.execute(OUTPUT_FILE..".htm")

if iup.Alarm("Open List of New Created Account", "Buka file daftar rekening baru yang dibuat \ndalam periode "..posisi_report1.." sampai "..posisi_report2.." ? " ,"Ya" ,"Tidak") == 1 then
	os.execute(OUTPUT_FILE.."_NEW.csv")
end


