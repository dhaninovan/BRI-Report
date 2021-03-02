-- Lua 5.1.5.XL Test Script for compress anyfile using gzio Library

-- stream the text file into a gzip file
params = {...}
filename = params[1]
gzFile = assert(gzio.open(filename..".gz", "w"))
for line in io.lines(filename) do
	gzFile:write(line..'\n')
end
gzFile:close()
iup.Message("Info", "Success compress "..filename.." into "..filename..".gz\nFile size reduced: "..(os.getfilesize(filename)-os.getfilesize(filename..'.gz')..' bytes'))
