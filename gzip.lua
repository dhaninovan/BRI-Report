-- Lua 5.1.5.XL Test Script for compress anyfile using gzio Library

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

-- stream the text file into a gzip file
params = {...}
filename = params[1]

t1 = os.clock()
fi = assert(io.open(filename, "r"))
gzFile = assert(gzio.open(filename..".gz", "w"))
repeat
	buf = fi:read(1024)
	if buf ~= nil then gzFile:write(buf) end
until buf == nil
gzFile:close()
fi:close()

iup.Message("Info", "Successfully compress "..filename.." into \n"..filename..".gz in "..(os.clock()-t1).."\nFile size reduced: "..format_number(os.getfilesize(filename)-os.getfilesize(filename..'.gz')..' bytes'))
