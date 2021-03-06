-- Diff Report : Find and calculate the difference between files DI319PN Saving Account Balance (CSV Format from Portal DWH)
-- Interpreter : (Novan's Modified)lua.exe
-- 7:57 PM Thursday, April 25, 2013 Raha City

--[[
TODO:
1. Tambahkan support untuk file input multi PN (Sesuaikan Field PN Pengelola)
2. Tambahkan field untuk Date Opened in Output New Account dan atur format sesuai dengan Indonesia [done]
3. Support multi currency

New Field in DI319 Multi PN:
25. PN_Customer_Service	
26. PN_RM_Dana	
27. PN_RM_Pinjaman	
28. PN_RM_Merchant	
29. PN_Relationship_Officer	
30. PN_Sales_Person	
31. PN_PAB	
32. PN_RM_Referral	
33. JUMLAH_PN

]]--

local list_acc = {}
local OUTPUT_FILE = "DI319PNDIFF"
local f_lines1 = nil
local f_lines2 = nil
local HEADER_DI319_MULTIPN = "textbox16,textbox8,textbox14,textbox22,textbox15,textbox4,textbox38,textbox6,DLT,textbox2,BALANCE,AVAILBALANCE,INTCREDIT,ACCRUEINT,AVRGBALANCE,textbox11,textbox18,textbox23,KECAMATAN_T_TINGGAL,KELURAHAN_T_TINGGAL,KODEPOS_T_TINGGAL,KECAMATAN_T_USAHA,KELURAHAN_T_USAHA,KODEPOS_T_USAHA,PN_Customer_Service,PN_RM_Dana,PN_RM_Pinjaman,PN_RM_Merchant,PN_Relationship_Officer,PN_Sales_Person,PN_PAB,PN_RM_Referral,JUMLAH_PN"
local HEADER_DI319_MULTIPN_V2 = "textbox16,textbox8,textbox14,textbox22,textbox15,textbox4,textbox38,DLT,textbox2,BALANCE,AVAILBALANCE,INTCREDIT,ACCRUEINT,AVRGBALANCE,textbox11,KECAMATAN_T_TINGGAL,KELURAHAN_T_TINGGAL,KODEPOS_T_TINGGAL,KECAMATAN_T_USAHA,KELURAHAN_T_USAHA,KODEPOS_T_USAHA,PN_Customer_Service,PN_RM_Dana,PN_RM_Pinjaman,PN_RM_Merchant,PN_Relationship_Officer,PN_Sales_Person,PN_PAB,PN_RM_Referral,JUMLAH_PN"
local HEADER_DI319_PN = 	 "textbox16,textbox8,textbox14,textbox22,textbox15,textbox4,textbox38,textbox6,DLT,textbox2,BALANCE,AVAILBALANCE,INTCREDIT,ACCRUEINT,AVRGBALANCE,textbox11,textbox18,textbox23,KECAMATAN_T_TINGGAL,KELURAHAN_T_TINGGAL,KODEPOS_T_TINGGAL,KECAMATAN_T_USAHA,KELURAHAN_T_USAHA,KODEPOS_T_USAHA"

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

function Rekap_Officer(cs, rm_dana, rm_kredit, rm_merchant, rm_ro, rm_sp, rm_pab, rm_referal)
	local res = {}
	
	if #cs ~= 0 then res[#res+1] = "CS = "..cs end
	if #rm_dana ~= 0 then res[#res+1] = "RM Dana = "..rm_dana end
	if #rm_kredit ~= 0 then res[#res+1] = "RM Kredit = "..rm_kredit end
	if #rm_merchant ~= 0 then res[#res+1] = "RM Merchant = "..rm_merchant end
	if #rm_ro ~= 0 then res[#res+1] = "RO = "..rm_ro end
	if #rm_sp ~= 0 then res[#res+1] = "SP = "..rm_sp end
	if #rm_pab ~= 0 then res[#res+1] = "PAB = "..rm_pab end
	if #rm_referal ~= 0 then res[#res+1] = "Referal = "..rm_referal end
	return table.concat(res,", ")
end

function Report_Type(header)
	header = header:gsub(string.char(0x0D),'')
	header = header:gsub(string.char(0x0A),'')
	if header == string.char(0xEF, 0xBB, 0xBF)..HEADER_DI319_MULTIPN then
		return "DI319MULTIPN"
	elseif header == string.char(0xEF, 0xBB, 0xBF)..HEADER_DI319_MULTIPN_V2 then
		return "DI319MULTIPN_V2"
	elseif header == string.char(0xEF, 0xBB, 0xBF)..HEADER_DI319_PN then
		return "DI319PN"
	elseif header == HEADER_DI319_MULTIPN then 
		return "DI319MULTIPN"
	elseif header == HEADER_DI319_MULTIPN_V2 then 
		return "DI319MULTIPN_V2"
	elseif header == HEADER_DI319_PN then
		return "DI319PN"
	else
		print(header)
		return nil
	end
end

res, dummy, ReportFileName1, ReportFileName2, limit_res = iup.GetParam("Pilih Report DI319 MULTI PN dalam Format CSV (Sumber: DWH)", nil, [=[
Sumber Data: %m\n
Report Posisi Awal: %f[OPEN|*DI319*.csv;*DI319*.gz|CURRENT|NO|NO]\n
Report Posisi Akhir: %f[OPEN|*DI319*.csv;*DI319*.gz|CURRENT|NO|NO]\n
Limit Result: %l|10|20|30|50|100|\n
]=]
,"1. Buka Aplikasi BRISIM (https://brisim.bri.co.id)\n2. Pilih: DWH Reports\n3. Pilih: Critical Report\n4. Pilih: Table\n5. Pilih DI319 - MULTI PN SAVINGS ACCOUNT MONTHLY TRIAL BALANCE - ACTIVE (1 ROW)\n6. Download dan Save dalam format CSV", "C:\\Lua\\data\\20210131 DI319 MULTI PN.csv.gz","C:\\Lua\\data\\20210225 DI319 MULTI PN.csv.gz",1)

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
f_lines1 = io.lines
else
f_lines1 = gzio.lines
end

if ReportFileName2:match('%.gz$') == nil then
print('Converting '..ReportFileName2..' to ANSI encoding')
os.execute('type "'..ReportFileName2..'" > '..'tmp.csv')
os.remove(ReportFileName2)
os.rename('tmp.csv', ReportFileName2)
f_lines2 = io.lines
else
f_lines2 = gzio.lines
end

-- Load first data into table list_acc
t1 = os.clock()
print('Loading data from '..ReportFileName1)
no = 1
sep = ','
posisi_report1 = ''
report1_type = ''
report2_type = ''
for line in f_lines1(ReportFileName1) do
	-- process header
	if no == 1 then
		sep = FindFirstSeparator(line)
		report1_type = Report_Type(line)
		if report1_type ~= "DI319PN" and report1_type ~= "DI319MULTIPN" and report1_type ~= "DI319MULTIPN_V2" then
			iup.Message("Error","Report [Awal] yang dipilih bukan Report \"DI319\" PN atau \"DI319 MULTI PN\" dalam format CSV.\nSilahkan download ulang dari BRISIM atau \npilih kembali report yang sesuai.")
			return -1
		end
	else
		-- process data
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			posisi_report1 = f[1]
			acc_no = f[5]
			acc_cif = f[6]
			acc_name = f[7]
			if report1_type == "DI319MULTIPN" then
				acc_officer = Rekap_Officer(f[25],f[26],f[27],f[28],f[29],f[30],f[31],f[32])
			elseif report1_type == "DI319MULTIPN_V2" then
				acc_officer = Rekap_Officer(f[22],f[23],f[24],f[25],f[26],f[27],f[28],f[29])
			else
				acc_officer = f[17]..'-'..f[18]
			end
			if report1_type == "DI319MULTIPN_V2" then
				acc_balance = string.gsub(string.sub(f[10], 1, #f[10]-3), ",", "")
			else
				acc_balance = string.gsub(string.sub(f[11], 1, #f[11]-3), ",", "")
			end
			if (acc_balance ~= nil) then 
				acc_balance = tonumber(acc_balance)
				if (acc_balance ~= nil) then 
					list_acc[acc_no] = {acc_cif, acc_name, acc_balance, 0, -acc_balance, acc_officer}
				else
					print('Error in '..acc_no..' '..acc_name..' Balance: '..string.gsub(string.sub(f[11], 1, #f[11]-3), ",", ""))
				end
			else
				print('Error in '..acc_no..' '..acc_name)
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
fo:write('Rekening'..output_sep..'Tipe'..output_sep..'Nama'..output_sep..'Tanggal Buka'..output_sep..'Saldo'..output_sep..'PN_Pengelola\n')
for line in f_lines2(ReportFileName2) do
	-- only process line begin with number, skipping header
	if no == 1 then
		sep = FindFirstSeparator(line)
		report2_type = Report_Type(line)
		if report2_type ~= "DI319PN" and report2_type ~= "DI319MULTIPN" and report2_type ~= "DI319MULTIPN_V2" then
			iup.Message("Error","Report [Akhir] yang dipilih bukan Report \"DI319\" PN atau \"DI319 MULTI PN\" dalam format CSV.\\nSilahkan download ulang dari BRISIM atau \\npilih kembali report yang sesuai.")
			return -1
		end
	else
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			posisi_report2 = f[1]
			acc_no = f[5]
			acc_cif = f[6]
			acc_name = f[7]
			if report2_type == "DI319MULTIPN" then
				acc_officer = Rekap_Officer(f[25],f[26],f[27],f[28],f[29],f[30],f[31],f[32])
			elseif report2_type == "DI319MULTIPN_V2" then
				acc_officer = Rekap_Officer(f[22],f[23],f[24],f[25],f[26],f[27],f[28],f[29])
			else
				acc_officer = f[17]..'-'..f[18]
			end
			if report2_type == "DI319MULTIPN_V2" then
				acc_balance = string.gsub(string.sub(f[10], 1, #f[10]-3), ",", "")
			else
				acc_balance = string.gsub(string.sub(f[11], 1, #f[11]-3), ",", "")
			end
			acc_balance = tonumber(acc_balance)
			if list_acc[acc_no] then
				list_acc[acc_no][4] = acc_balance
				list_acc[acc_no][5] = list_acc[acc_no][4] - list_acc[acc_no][3]
				list_acc[acc_no][6] = acc_officer
			else
				list_acc[acc_no] = {acc_cif, acc_name, 0, acc_balance, acc_balance, acc_officer}
				if report2_type == "DI319MULTIPN_V2" then
					dd = csv.parse(f[9],'/')
					fo:write(string.format('%s%s%s%s"%s"%s%s/%s/%s%s"%s"%s%s\n', 
					format_account(acc_no), output_sep,
					f[15], output_sep,
					acc_name, output_sep,
					dd[2], dd[1], dd[3], output_sep,
					format_number(acc_balance), output_sep,
					acc_officer))
				else
					dd = csv.parse(f[10],'/')
					fo:write(string.format('%s%s%s%s"%s"%s%s/%s/%s%s"%s"%s%s\n', 
					format_account(acc_no), output_sep,
					f[16], output_sep,
					acc_name, output_sep,
					dd[2], dd[1], dd[3], output_sep,
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
 
<title>BRI Reporting Tool: Delta Tabungan DI319</title>
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
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td align='right'>"..format_number(v[4]).."</td><td align='right'>"..format_number(v[5]).."</td><td class='minus' align='right'>"..format_number(v[6]).."</td><td align='right'>"..v[7].."</td><td/></tr>\n")
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
		fo2:write("<tr><td>"..format_account(v[1]).."</td><td>"..v[2].."</td><td>"..v[3].."</td><td align='right'>"..format_number(v[4]).."</td><td align='right'>"..format_number(v[5]).."</td><td class='plus' align='right'>"..format_number(v[6]).."</td><td align='right'>"..v[7].."</td><td/></tr>\n")
	end
	
	fo:write(v[1]..output_sep..v[2]..output_sep..'"'..v[3]..'"'..output_sep..v[4]..output_sep..v[5]..output_sep..v[6]..output_sep..v[7]..'\n')
	ii = ii + 1
end
fo2:write([=[
	</tbody>
 </table>
 &nbsp;&nbsp;&nbsp;&nbsp;Data selengkapnya: <a href=']=]..OUTPUT_FILE..[=[.csv'>]=]..OUTPUT_FILE..[=[.csv</a>
 <hr>
 <div align='center' style='font-size:smaller'>BRI Reporting Tool: DI319 Diff<br>Copyright &copy;2013, <b>Dhani Novan</b> (dhani_novan@bri.co.id)</div><br/>
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
--os.execute("pause")
