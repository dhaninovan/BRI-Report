-- Display Performance each funding officer based on DI319PN Saving Account Balance (CSV Format from DWH BRISIM)
-- There is bug in report DI319PN. There are double data. So we have to remove duplicate date. We can use Notepad2 : Edit -> Block -> Sort Lines : Merge Duplicate Line
-- Interpreter : (Novan's Modified)lua.exe
-- 5:29 07 September 2019, Rawamangun

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
list_summ = {}
list_RM = {}
-- 0. PN RM PENGELOLA : key
-- 1. NAMA RM PENGELOLA
OUTPUT_FILE = "Tabungan_Performance_Officer_Report"
output_sep = ";"

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

res, ReportFileName1, ReportFileName2 = iup.GetParam("Pick Reports DI319PN in CSV Format (Source: BRISIM)", nil, [=[
Current Report  : %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Baseline Report : %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
]=]
,"","")

if ReportFileName1 == "" or ReportFileName2 == "" then
	print("Please select two reports to be compared")
	os.exit(-1)
end
if ReportFileName1 == nil or ReportFileName2 == nil then
	print("Please select two reports to be compared")
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

-- Load first data (current report) into table list_acc
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
			acc_no = f[5]
			acc_cif = f[6]
			acc_name = f[7]
			acc_balance = string.gsub(string.sub(f[11], 1, #f[11]-3), ",", "")
			acc_balance = tonumber(acc_balance)
			acc_officer = f[17]
--			1	CIF
--			2	Name
--			3	Account Baseline Balance
--			4	Account Current Balance
--			5	Delta Balance
--			6	Account Officer
--			7	Remark
			list_RM[f[17]] = f[18]
			if f[3] == "IDR " then
				list_acc[acc_no] = {acc_cif, acc_name, 0, acc_balance, acc_balance, acc_officer, "Tabungan baru"}
			end
			
			-- update Officer name in summary table from f[18]
			if list_summ[acc_officer] == nil then
				list_summ[acc_officer] = {f[18], 0, 0, 0, 0}
			end
		end
	end
	no = no + 1
end

-- Update table list_acc with baseline data
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
			acc_no = f[5]
			acc_cif = f[6]
			acc_name = f[7]
			acc_balance = string.gsub(string.sub(f[11], 1, #f[11]-3), ",", "")			
			acc_balance = tonumber(acc_balance)
			list_RM[f[17]] = f[18]
			if f[3] == "IDR " then
			if list_acc[acc_no] then
				list_acc[acc_no][3] = acc_balance
				list_acc[acc_no][5] = list_acc[acc_no][4] - list_acc[acc_no][3]
				acc_officer = list_acc[acc_no][6]
				
				-- update remark based on old and new PN
				if f[17] == list_acc[acc_no][6] then
					-- update remark based on old and new balance
					if list_acc[acc_no][5] == 0 then
						list_acc[acc_no][7] = ""
					elseif list_acc[acc_no][5] > 0 then
						list_acc[acc_no][7] = "Top Up"
					elseif list_acc[acc_no][5] < 0 then
						list_acc[acc_no][7] = "Penurunan Saldo"
					end
				elseif f[17] == "" then
					if list_acc[acc_no][5] > 0 then
						list_acc[acc_no][7] = string.format("Patching tabungan tanpa PN menjadi PN %s(%s). Saldo meningkat", list_acc[acc_no][6], list_RM[list_acc[acc_no][6]])
					elseif list_acc[acc_no][5] < 0 then
						list_acc[acc_no][7] = string.format("Patching tabungan tanpa PN menjadi PN %s(%s). Saldo menurun", list_acc[acc_no][6], list_RM[list_acc[acc_no][6]])
					end
				else
					if list_acc[acc_no][5] > 0 then
						list_acc[acc_no][7] = string.format("Patching tabungan PN %s(%s) menjadi PN %s(%s). Saldo meningkat", f[17], list_RM[f[17]], list_acc[acc_no][6], list_RM[list_acc[acc_no][6]])
					elseif list_acc[acc_no][5] < 0 then
						list_acc[acc_no][7] = string.format("Patching tabungan PN %s(%s) menjadi PN %s(%s). Saldo menurun", f[17], list_RM[f[17]], list_acc[acc_no][6], list_RM[list_acc[acc_no][6]])
					end
				end
			else
				acc_officer = f[17]
				list_acc[acc_no] = {acc_cif, acc_name, acc_balance, 0, -acc_balance, acc_officer, "Penutupan Tabungan"}
			end
			end
			-- update Officer name in summary table from f[18]
			if list_summ[acc_officer] == nil then
				list_summ[acc_officer] = {f[18], 0, 0, 0, 0}
			end
		end
	end
	no = no + 1
end

print('Sorting descending')
sorted_list_acc = {}
for k, v in pairs(list_acc) do
	table.insert(sorted_list_acc, {k, v[1], v[2], v[3], v[4], v[5], v[6], v[7]})
	
	-- update summary (baseline balance) for each officer
	acc_officer = v[6]
	if v[3] > 0 then
		list_summ[acc_officer][2] = list_summ[acc_officer][2] + 1
		list_summ[acc_officer][3] = list_summ[acc_officer][3] + v[3]
	end
	
	-- update summary (current balance) for each officer
	if v[4] > 0 then
		list_summ[acc_officer][4] = list_summ[acc_officer][4] + 1
		list_summ[acc_officer][5] = list_summ[acc_officer][5] + v[4]
	end
end
table.sort(sorted_list_acc, function(a,b) return a[6]<b[6] end)
print("Processing "..tostring(#sorted_list_acc).." rekening")

print('Writing '..OUTPUT_FILE)
if output_sep == 0 then
	output_sep = ','
elseif output_sep == 1 then
	output_sep = ';'
end

fo2 = io.open(OUTPUT_FILE..".htm", "w")
fo2:write([=[
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 
<title>BRI Reporting Tool: Tabungan DI319PN</title>
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
 ]=])
 
 fo2:write([=[ 
 <h2>&nbsp;&nbsp;&nbsp;Summary Tabungan Performance Report (Periode ]=]..posisi_report2..[=[ - ]=]..posisi_report1..[=[)</h2>
 <table cellspacing='0'>
	<thead>
		<tr>
			<th>Personal No</th>
			<th>Account Officer Name</th>
			<th>Count of baseline account<br>]=]..posisi_report2..[=[</th>
			<th>Sum of baseline account balance<br>]=]..posisi_report2..[=[</th>
			<th>Count of current account<br>]=]..posisi_report1..[=[</th>
			<th>Sum of current account balance<br>]=]..posisi_report1..[=[</th>
		</tr>
	</thead>
	<tbody>
]=])
for k, v in pairs(list_summ) do
	fo2:write("<tr><td>"..k.."</td><td>"..v[1].."</td><td>"..format_number(v[2]).."</td><td align='right'>"..format_number(v[3]).."</td><td align='right'>"..format_number(v[4]).."</td><td align='right'>"..format_number(v[5]).."</td></tr>\n")
end
fo2:write([=[
	</tbody>
 </table>
]=])
fo2:write([=[
 <hr>
 <div align='center' style='font-size:smaller'>BRI Reporting Tool<br>Copyright &copy;2013, <b>Dhani Novan</b> (dhani_novan@bri.co.id)</div><br/>
</body>
</html>
]=])
fo2:close()

fo2 = io.open(OUTPUT_FILE..".csv", "w")
fo2:write('Account No'..output_sep..'CIF'..output_sep..'Account Name'..output_sep..'Base Balance'..output_sep..'Current Balance'..output_sep..'Delta Balance'..output_sep..'PN Officer'..output_sep..'Remark\n')
for k, v in pairs(sorted_list_acc) do
	fo2:write(v[1]..output_sep..v[2]..output_sep..'"'..v[3]..'"'..output_sep..v[4]..output_sep..v[5]..output_sep..v[6]..output_sep..v[7]..output_sep..v[8]..'\n')
end
fo2:close()

print('=== Done in '..(os.clock()-t1)..' ===')
os.execute(OUTPUT_FILE..".htm")
os.execute("pause")
