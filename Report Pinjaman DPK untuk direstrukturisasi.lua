-- Report untuk menampilkan pinjaman DPK setiap PN Officer dengan urutan DPK3, DPK2, dan DPK1
-- Input:
--	1. LW321PNSingleRow
--	2. LW321PNSingleRow_Impact_Corona_V2
-- Interpreter : (Novan's Modified)lua.exe v1.1.1.0
-- 5:05 02 June 2020 Rawamangun

-- TODO : DIFF local tm = os.time{year=1999, month=12, day=01}

--	LW321PNSingleRow
--  1.Periode
--  2.Branch
--  3.Currency
--  4.Nama AO
--  5.LN Type
--  6.Nomor rekening
--  7.Nama Debitur
--  8.Plafond
--  9.Next Pmt Date
-- 10.Next Int Pmt Date
-- 11.Rate
-- 12.Tgl Menunggak
-- 13.Tgl Realisasi
-- 14.Tgl Jatuh tempo
-- 15.Jangka Waktu
-- 16.Flag Restruk
-- 17.CIFNO
-- 18.Kolektibilitas Lancar
-- 19.Kolektibilitas DPK
-- 20.Kolektibilitas Kurang Lancar
-- 21.Kolektibilitas Diragukan
-- 22.Kolektibilitas Macet
-- 23.Tunggakan Pokok
-- 24.Tunggakan Bunga
-- 25.Tunggakan Pinalty
-- 26.PN
-- 27.Nama PN
-- 28.Code
-- 29.Description
-- 30.Kol_ADK
	

list_acc = {}
OUTPUT_FILE = "Report_DPK"

-- convert string to date
-- format input string: 03/06/2020 OR 03-06-2020 OR 3/6/2020 OR 3-6-2020
function todate(idate)
	local dd, mm, yyyy = string.match(idate,'(%d*%d)%D(%d*%d)%D(%d%d%d%d)')
	dd = tonumber(dd)
	mm = tonumber(mm)
	yyyy = tonumber(yyyy)
	
	return os.time{day=dd, month=mm, year=yyyy}
end

-- Date1(string) : Next Pmt Date
-- Date2(string) : Int Pmt Date
function hitung_umur_tunggakan(date1, date2)
	if #date1 ~= 0 and #date2 ~= 0 then
		local dt1 = todate(date1)
		local dt2 = todate(date2)
		if  dt1 < dt2 then
			return math.ceil((os.time()-dt1) / (3600*24))
		else
			return math.ceil((os.time()-dt2) / (3600*24))
		end
	elseif #date1 ~= 0 and #date2 == 0 then
		return math.ceil((os.time()-todate(date1)) / (3600*24))
	elseif #date1 == 0 and #date2 ~= 0 then
		return math.ceil((os.time()-todate(date2)) / (3600*24))
	elseif #date1 == 0 and #date2 == 0 then
		return 0
	end
end

function format_account(s)
    local rek_len = string.len(s)
    if rek_len >= 12 and rek_len <=19 then
        s = string.rep("0",15-rek_len)..s
        s = table.concat({string.sub(s,1,-12), string.sub(s,-11,-10), string.sub(s,-9,-4), string.sub(s,-3,-2), string.sub(s,-1,-1)}, '-')
    end
    return s
end

function format_number(v)
	local s
	local unary
	
	if v < 0 then
		s = tostring(-v)
		unary = "-"
	else
		s = tostring(v)
		unary = ""
	end
    
    local pos = string.len(s) % 3

    if pos == 0 then pos = 3 end
    return unary..string.sub(s, 1, pos).. string.gsub(string.sub(s, pos+1), "(...)", ".%1")
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

res, ReportFileName1, ReportFileName2, output_sep = iup.GetParam("Pilih Report Pinjaman LW321 dalam Format CSV (Sumber: DWH)", nil, [=[
Report LW321PNSingleRow_Impact_Corona_V2: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Report LW321PNSingleRow: %f[OPEN|*.csv;*.txt|CURRENT|NO|NO]\n
Output Separator: %l|,|;|\n
]=]
--,"E:\\Projects\\LUA Script\\BRI Report\\Data\\LW321PNSingleRow%5FImpact%5FCorona%5FV2 KC.csv","E:\\Projects\\LUA Script\\BRI Report\\Data\\LW321PNSingleRow KC.csv",0)
,"E:\\Projects\\LUA Script\\BRI Report\\Data\\LW321PNSingleRow%5FImpact%5FCorona%5FV2 ALLKCP.csv","E:\\Projects\\LUA Script\\BRI Report\\Data\\LW321PNSingleRow ALLKCP.csv",0)
--,"D:\\Data Pinjaman\\20200531\\LW321PNSingleRow%5FImpact%5FCorona%5FV2 KC.csv","D:\\Data Pinjaman\\20200531\\LW321PNSingleRow KC.csv",0)

if ReportFileName1 == "" or ReportFileName2 == "" then
	print("Please select two reports to be compared")
	os.execute("pause")
	os.exit(-1)
end

if output_sep == 0 then
	output_sep = ','
elseif output_sep == 1 then
	output_sep = ';'
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

-- Load first data into string : rek_restruk
t1 = os.clock()
print('Loading data from '..ReportFileName1)
no = 1
sep = ','
local t_tmp = {}
for line in io.lines(ReportFileName1) do
	-- process header
	if no == 1 then
		sep = FindFirstSeparator(line)
	else
		-- process data
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			t_tmp[#t_tmp + 1] = f[6]
		end
	end
	no = no + 1
end
rek_restruk = table.concat(t_tmp, " ")

-- Cek each data in LW321PNSingleRow
print('Loading data from '..ReportFileName2)
no = 1
sep = ','
list_acc = {}
for line in io.lines(ReportFileName2) do
	-- only process line begin with number, skipping header
	if no == 1 then
		sep = FindFirstSeparator(line)
	else
		f = csv.parse(line, sep)
		if f[1] ~= "" then
			acc_no = f[6]
			acc_balance = f[19]	-- OS kolek DPK
			acc_balance = string.gsub(string.sub(acc_balance, 1, #acc_balance-3), ",", "")
			acc_balance = tonumber(acc_balance)
			if string.find(rek_restruk, acc_no) == nil and acc_balance ~= 0 then
				--acc_type = f[5]
				--acc_name = f[7]
				--acc_officer = f[27]
				list_acc[#list_acc+1] = {acc_no, f[7], f[5], acc_balance, hitung_umur_tunggakan(f[9], f[10]), f[27], f[2]}
			end
		end
	end
	no = no + 1
end

-- sort by uker(output field no 7), officer(output field no 6), umur tunggakan (output field no 5)
function custom_sort(a,b)
	if a[7] == b[7] then	
		if a[6] == b[6] then
			return a[5] > b[5]	--compare field 'umur tunggakan'
		else
			return a[6] < b[6]	--compare field 'officer'
		end
	else
		return a[7] < b[7]	--compare field 'uker'
	end
end


table.sort(list_acc, custom_sort) 
fo = io.open(OUTPUT_FILE..".csv", "w")
fo:write(table.concat({'Rekening', 'Nama Debitur', 'Tipe', 'Outstanding', 'Tunggakan(hari)', 'Officer', 'Uker'}, output_sep)..'\n')
for k, v in pairs(list_acc) do
	fo:write(table.concat({format_account(v[1]), v[2], v[3], format_number(v[4]), v[5], v[6], v[7]}, output_sep)..'\n')
end
fo:close()
print('done')
os.execute(OUTPUT_FILE..".csv")
os.execute("pause")
