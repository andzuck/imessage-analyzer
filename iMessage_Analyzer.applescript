-- Paste into macOS Script Editor, then press Run.
-- It creates two CSV files and three browser-ready charts in a folder you choose:
--   1. iMessage_topn_halfyear_lines.csv Ñ chart-ready values for the top N per half-year
--   2. iMessage_halfyear_leaderboard.csv Ñ the top N independently ranked within every half-year
--
-- Counts only one-to-one conversations (not group threads), and includes sent + received iMessages.
use scripting additions
property appleEpochOffset : 9.783072E+8
set currentYear to (do shell script "/bin/date +%Y") as integer
set startAnswer to text returned of (display dialog "First year to include:" default answer "2015" buttons {"Cancel", "Continue"} default button "Continue")
set endAnswer to text returned of (display dialog "Last year to include:" default answer (currentYear as text) buttons {"Cancel", "Continue"} default button "Continue")
try
	set startYear to startAnswer as integer
	set endYear to endAnswer as integer
on error
	display alert "Please enter whole years, such as 2018 and 2026."
	return
end try
if startYear > endYear then
	display alert "The first year must be no later than the last year."
	return
end if
set topNAnswer to text returned of (display dialog "How many contacts should each chart include?" default answer "10" buttons {"Cancel", "Continue"} default button "Continue")
try
	set topN to topNAnswer as integer
on error
	display alert "Please enter a whole number, such as 5 or 10."
	return
end try
if topN < 1 then
	display alert "The number of contacts must be at least 1."
	return
end if
set exportFolder to choose folder with prompt "Choose a folder for the CSV exports:"
set dbPath to POSIX path of ((path to library folder from user domain as text) & "Messages:chat.db")
set linesPath to POSIX path of exportFolder & "iMessage_topn_halfyear_lines.csv"
set leadersPath to POSIX path of exportFolder & "iMessage_halfyear_leaderboard.csv"
set chartDataPath to POSIX path of exportFolder & "iMessage_topn_halfyear_lines.json"
set stackedDataPath to POSIX path of exportFolder & "iMessage_topn_halfyear_sent_received.json"
set allTimeMessagesPath to POSIX path of exportFolder & "iMessage_topn_alltime_messages.json"
set allTimeDaysPath to POSIX path of exportFolder & "iMessage_topn_alltime_days.json"
set dashboardPath to POSIX path of exportFolder & "iMessage_contact_analysis.html"
-- Finder paths with apostrophes are legal, so quote those safely for SQLite.
set dbSQLPath to my sqlQuote(dbPath)
set linesSQLPath to my sqlQuote(linesPath)
set leadersSQLPath to my sqlQuote(leadersPath)
set chartDataSQLPath to my sqlQuote(chartDataPath)
set stackedDataSQLPath to my sqlQuote(stackedDataPath)
set allTimeMessagesSQLPath to my sqlQuote(allTimeMessagesPath)
set allTimeDaysSQLPath to my sqlQuote(allTimeDaysPath)
set dateExpression to "datetime((CASE WHEN m.date > 100000000000 THEN m.date / 1000000000 ELSE m.date END) + " & appleEpochOffset & ", 'unixepoch', 'localtime')"
try
	set contactMapJSON to my contactNameMapJSON()
on error
	-- The chart remains usable with phone/email handles if Contacts access is declined.
	set contactMapJSON to "{}"
end try
-- Each half-year gets its own top-N list. Chart lines deliberately break when a
-- contact is not in that period's list, avoiding a misleading continuous trend.
set commonCTE to "WITH RECURSIVE" & linefeed & Â
	"	one_to_one_chats AS (" & linefeed & Â
	"		SELECT chat_id, MIN(handle_id) AS handle_id" & linefeed & Â
	"		FROM chat_handle_join GROUP BY chat_id HAVING COUNT(DISTINCT handle_id) = 1" & linefeed & Â
	"	)," & linefeed & Â
	"	halfyearly AS (" & linefeed & Â
	"		SELECT CAST(strftime('%Y', " & dateExpression & ") AS INTEGER) AS year," & linefeed & Â
	"		       CASE WHEN CAST(strftime('%m', " & dateExpression & ") AS INTEGER) <= 6 THEN 1 ELSE 2 END AS half," & linefeed & Â
	"		       h.id AS contact, COUNT(DISTINCT m.ROWID) AS message_count," & linefeed & Â
	"		       SUM(CASE WHEN m.is_from_me = 1 THEN 1 ELSE 0 END) AS sent_count," & linefeed & Â
	"		       SUM(CASE WHEN m.is_from_me = 1 THEN 0 ELSE 1 END) AS received_count" & linefeed & Â
	"		FROM message m" & linefeed & Â
	"		JOIN chat_message_join cmj ON cmj.message_id = m.ROWID" & linefeed & Â
	"		JOIN one_to_one_chats o ON o.chat_id = cmj.chat_id" & linefeed & Â
	"		JOIN handle h ON h.ROWID = o.handle_id" & linefeed & Â
	"		WHERE m.date > 0 AND h.id IS NOT NULL" & linefeed & Â
	"		  AND (m.associated_message_type IS NULL OR m.associated_message_type = 0)" & linefeed & Â
	"		GROUP BY year, half, h.id" & linefeed & Â
	"	)," & linefeed
set oldLinesQuery to commonCTE & Â
	"	halves(year, half) AS (" & linefeed & Â
	"		SELECT " & startYear & ", 1" & linefeed & Â
	"		UNION ALL SELECT CASE WHEN half = 2 THEN year + 1 ELSE year END," & linefeed & Â
	"		                 CASE WHEN half = 2 THEN 1 ELSE half + 1 END" & linefeed & Â
	"		FROM halves WHERE year < " & endYear & " OR half < 2" & linefeed & Â
	"	)," & linefeed & Â
	"	top_contacts AS (" & linefeed & Â
	"		SELECT contact, SUM(message_count) AS period_total FROM halfyearly" & linefeed & Â
	"		WHERE year BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"		GROUP BY contact ORDER BY period_total DESC, contact LIMIT " & topN & linefeed & Â
	"	)" & linefeed & Â
	"SELECT q.year, q.half, printf('%d-H%d', q.year, q.half) AS season," & linefeed & Â
	"	       t.contact, COALESCE(x.message_count, 0) AS message_count, t.period_total" & linefeed & Â
	"FROM halves q CROSS JOIN top_contacts t" & linefeed & Â
	"LEFT JOIN halfyearly x ON x.year = q.year AND x.half = q.half AND x.contact = t.contact" & linefeed & Â
	"ORDER BY q.year, q.half, t.period_total DESC, t.contact;"
set linesQuery to commonCTE & Â
	"	ranked AS (" & linefeed & Â
	"		SELECT year, half, contact, message_count," & linefeed & Â
	"		       ROW_NUMBER() OVER (PARTITION BY year, half ORDER BY message_count DESC, contact) AS rank" & linefeed & Â
	"		FROM halfyearly WHERE year BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"	)" & linefeed & Â
	"SELECT year, half, printf('%d-H%d', year, half) AS season, contact, message_count, rank" & linefeed & Â
	" FROM ranked WHERE rank <= " & topN & " ORDER BY year, half, rank;"
-- This companion view is useful for answering the literal "who was #1 that season?"
-- question; contacts may change from one half-year to the next.
set leadersQuery to commonCTE & Â
	"	ranked AS (" & linefeed & Â
	"		SELECT year, half, contact, message_count," & linefeed & Â
	"		       ROW_NUMBER() OVER (PARTITION BY year, half ORDER BY message_count DESC, contact) AS rank" & linefeed & Â
	"		FROM halfyearly WHERE year BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"	)" & linefeed & Â
	"SELECT year, half, printf('%d-H%d', year, half) AS season, rank, contact, message_count" & linefeed & Â
	" FROM ranked WHERE rank <= " & topN & " ORDER BY year, half, rank;"
set stackedQuery to commonCTE & Â
	"	halves(year, half) AS (" & linefeed & Â
	"		SELECT " & startYear & ", 1" & linefeed & Â
	"		UNION ALL SELECT CASE WHEN half = 2 THEN year + 1 ELSE year END," & linefeed & Â
	"		                 CASE WHEN half = 2 THEN 1 ELSE half + 1 END" & linefeed & Â
	"		FROM halves WHERE year < " & endYear & " OR half < 2" & linefeed & Â
	"	)," & linefeed & Â
	"	top_contacts AS (" & linefeed & Â
	"		SELECT contact, SUM(message_count) AS period_total FROM halfyearly" & linefeed & Â
	"		WHERE year BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"		GROUP BY contact ORDER BY period_total DESC, contact LIMIT " & topN & linefeed & Â
	"	)" & linefeed & Â
	"SELECT q.year, q.half, printf('%d-H%d', q.year, q.half) AS season, t.contact," & linefeed & Â
	"       COALESCE(x.sent_count, 0) AS sent_count, COALESCE(x.received_count, 0) AS received_count," & linefeed & Â
	"       COALESCE(x.message_count, 0) AS message_count, t.period_total" & linefeed & Â
	"FROM halves q CROSS JOIN top_contacts t" & linefeed & Â
	"LEFT JOIN halfyearly x ON x.year = q.year AND x.half = q.half AND x.contact = t.contact" & linefeed & Â
	"ORDER BY q.year, q.half, t.period_total DESC, t.contact;"
set allTimeMessagesQuery to commonCTE & Â
	"	totals AS (" & linefeed & Â
	"		SELECT contact, SUM(message_count) AS message_count FROM halfyearly" & linefeed & Â
	"		WHERE year BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"		GROUP BY contact" & linefeed & Â
	"	)" & linefeed & Â
	"SELECT contact, message_count FROM totals ORDER BY message_count DESC, contact LIMIT " & topN & ";"
set allTimeDaysQuery to "WITH one_to_one_chats AS (" & linefeed & Â
	"  SELECT chat_id, MIN(handle_id) AS handle_id FROM chat_handle_join" & linefeed & Â
	"  GROUP BY chat_id HAVING COUNT(DISTINCT handle_id) = 1" & linefeed & Â
	")" & linefeed & Â
	"SELECT h.id AS contact, COUNT(DISTINCT date(" & dateExpression & ")) AS day_count" & linefeed & Â
	" FROM message m JOIN chat_message_join cmj ON cmj.message_id=m.ROWID" & linefeed & Â
	" JOIN one_to_one_chats o ON o.chat_id=cmj.chat_id JOIN handle h ON h.ROWID=o.handle_id" & linefeed & Â
	" WHERE m.date > 0 AND h.id IS NOT NULL" & linefeed & Â
	"   AND CAST(strftime('%Y', " & dateExpression & ") AS INTEGER) BETWEEN " & startYear & " AND " & endYear & linefeed & Â
	"   AND (m.associated_message_type IS NULL OR m.associated_message_type = 0)" & linefeed & Â
	" GROUP BY h.id ORDER BY day_count DESC, h.id LIMIT " & topN & ";"
set sqlFile to POSIX path of ((path to temporary items from user domain as text) & "imessage-halfyear-export.sql")
set sqlText to ".headers on" & linefeed & ".mode csv" & linefeed & ".output " & linesSQLPath & linefeed & linesQuery & linefeed & ".output " & leadersSQLPath & linefeed & leadersQuery & linefeed & ".mode json" & linefeed & ".output " & chartDataSQLPath & linefeed & linesQuery & linefeed & ".output " & stackedDataSQLPath & linefeed & stackedQuery & linefeed & ".output " & allTimeMessagesSQLPath & linefeed & allTimeMessagesQuery & linefeed & ".output " & allTimeDaysSQLPath & linefeed & allTimeDaysQuery & linefeed & ".quit" & linefeed
try
	my writeText(sqlText, sqlFile)
	do shell script "/usr/bin/sqlite3 -readonly " & quoted form of dbPath & " < " & quoted form of sqlFile
on error errMsg number errNum
	display alert "The export could not run." message "Script Editor needs Full Disk Access to read ~/Library/Messages/chat.db. Also confirm that /usr/bin/sqlite3 is installed.

" & errMsg as critical
	return
end try
set chartJSON to my readText(chartDataPath)
set stackedJSON to my readText(stackedDataPath)
set allTimeMessagesJSON to my readText(allTimeMessagesPath)
set allTimeDaysJSON to my readText(allTimeDaysPath)
my writeText(my makeDashboardHTML(chartJSON, stackedJSON, allTimeMessagesJSON, allTimeDaysJSON, contactMapJSON, topN, startYear, endYear), dashboardPath)
do shell script "/usr/bin/open " & quoted form of dashboardPath
display dialog "Done. Created the CSV exports and opened the complete contact-analysis dashboard in your browser." buttons {"OK"} default button "OK"
on contactNameMapJSON()
	set mapEntries to {}
	tell application "Contacts"
		set everyone to every person
		repeat with personRecord in everyone
			set personName to name of personRecord as text
			repeat with emailRecord in emails of personRecord
				set addressValue to value of emailRecord as text
				set end of mapEntries to my jsonQuote(my normalizedContactID(addressValue)) & ":" & my jsonQuote(personName)
			end repeat
			repeat with phoneRecord in phones of personRecord
				set phoneValue to value of phoneRecord as text
				set end of mapEntries to my jsonQuote(my normalizedContactID(phoneValue)) & ":" & my jsonQuote(personName)
			end repeat
		end repeat
	end tell
	set AppleScript's text item delimiters to ","
	set mapJSON to "{" & (mapEntries as text) & "}"
	set AppleScript's text item delimiters to ""
	return mapJSON
end contactNameMapJSON
on normalizedContactID(rawID)
	if rawID contains "@" then return my lowerText(rawID)
	set allowedCharacters to "0123456789"
	set normalizedID to ""
	repeat with currentCharacter in characters of rawID
		if allowedCharacters contains (currentCharacter as text) then set normalizedID to normalizedID & currentCharacter
	end repeat
	return normalizedID
end normalizedContactID
on lowerText(sourceText)
	set uppercaseLetters to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	set lowercaseLetters to "abcdefghijklmnopqrstuvwxyz"
	set resultText to ""
	repeat with currentCharacter in characters of sourceText
		set characterIndex to offset of (currentCharacter as text) in uppercaseLetters
		if characterIndex is 0 then
			set resultText to resultText & currentCharacter
		else
			set resultText to resultText & character characterIndex of lowercaseLetters
		end if
	end repeat
	return resultText
end lowerText
on jsonQuote(sourceText)
	set escapedText to my replaceText(sourceText, "\\", "\\\\")
	set escapedText to my replaceText(escapedText, "\"", "\\\"")
	set escapedText to my replaceText(escapedText, linefeed, "\\n")
	set escapedText to my replaceText(escapedText, return, "\\n")
	return "\"" & escapedText & "\""
end jsonQuote
on replaceText(sourceText, findText, replacementText)
	set AppleScript's text item delimiters to findText
	set textItems to text items of sourceText
	set AppleScript's text item delimiters to replacementText
	set resultText to textItems as text
	set AppleScript's text item delimiters to ""
	return resultText
end replaceText
on readText(filePath)
	set fileRef to open for access (POSIX file filePath)
	try
		set theText to read fileRef as Çclass utf8È
		close access fileRef
		return theText
	on error errMsg number errNum
		try
			close access fileRef
		end try
		error errMsg number errNum
	end try
end readText
on makeDashboardHTML(lineJSON, stackedJSON, messageJSON, daysJSON, contactMapJSON, topN, startYear, endYear)
	return "<!doctype html><meta charset=\"utf-8\"><title>iMessage contact analysis</title>" & linefeed & Â
		"<style>body{margin:28px;max-width:1400px;font:14px -apple-system,BlinkMacSystemFont,sans-serif;color:#172033;background:#fafafa}h1{font-size:26px}h2{margin:38px 0 12px;font-size:19px}.sub{color:#586174}.tables{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:18px}.panel{background:#fff;border:1px solid #d8dee8;border-radius:10px;padding:16px}table{width:100%;border-collapse:collapse}td,th{padding:7px;border-bottom:1px solid #e5e9f0;text-align:left}th:last-child,td:last-child{text-align:right}#charts{display:grid;gap:30px}.chart{overflow-x:auto;background:#fff;border:1px solid #d8dee8;border-radius:10px;padding:12px}.chart h3{margin:0 0 8px;font-size:16px}svg{min-width:900px}.grid{stroke:#d8dee8}.axis{fill:#586174;font-size:11px}.legend{font-size:12px}.legend-list{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px 18px;padding:12px 6px 2px}.legend-item{display:flex;gap:7px;min-width:0;align-items:flex-start}.swatch{width:16px;height:3px;margin-top:7px;flex:0 0 auto}.swatch-pair{display:flex;gap:3px;flex:0 0 auto;margin-top:7px}.legend-copy{min-width:0;overflow-wrap:anywhere;line-height:1.35}.legend-detail{color:#586174}.line{fill:none;stroke-width:2.5}.dot{stroke:#fff;stroke-width:1.5}.area{stroke:#fff;stroke-width:.7}</style>" & linefeed & Â
		"<h1>iMessage contact analysis</h1><p class=\"sub\">" & startYear & "Ð" & endYear & " á top " & topN & " contacts</p><div class=\"tables\"><section class=\"panel\"><h2>Most messages</h2><div id=\"messages\"></div></section><section class=\"panel\"><h2>Most days texted</h2><div id=\"days\"></div></section></div><div id=\"charts\"></div>" & linefeed & Â
		"<script>const lineRaw=" & lineJSON & ",stackRaw=" & stackedJSON & ",msgRaw=" & messageJSON & ",daysRaw=" & daysJSON & ",names=" & contactMapJSON & ",N=" & topN & ";const norm=v=>v.includes('@')?v.toLowerCase():v.replace(/[^0-9]/g,''),name=v=>names[norm(v)]||v,lines=lineRaw.map(r=>({...r,contact:name(r.contact)})),stack=stackRaw.map(r=>({...r,contact:name(r.contact)})),msgs=msgRaw.map(r=>({...r,contact:name(r.contact)})),days=daysRaw.map(r=>({...r,contact:name(r.contact)}));const esc=v=>String(v).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));const table=(rows,key,label)=>`<table><thead><tr><th>Contact</th><th>${label}</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.contact)}</td><td>${Number(r[key]).toLocaleString()}</td></tr>`).join('')}</tbody></table>`;document.getElementById('messages').innerHTML=table(msgs,'message_count','Messages');document.getElementById('days').innerHTML=table(days,'day_count','Days');" & linefeed & Â
		"const colors=i=>`hsl(${(i*137.508)%360} 68% 42%)`,ord=n=>n+(n%10===1&&n%100!==11?'st':n%10===2&&n%100!==12?'nd':n%10===3&&n%100!==13?'rd':'th'),detail=(c,m)=>{const a=lines.filter(r=>r.contact===c);if(m==='volume'){const t={};a.forEach(r=>t[r.year]=(t[r.year]||0)+r.message_count);return Object.entries(t).map(([y,n])=>`${y} Ð ${n.toLocaleString()}`).join('; ')}return a.map(r=>`${r.season} Ð ${ord(r.rank)}`).join('; ')},legend=(cs,m)=>`<div class='legend-list'>${cs.map((c,k)=>`<div class='legend-item'><span class='swatch' style='background:${colors(k)}'></span><div class='legend-copy'><strong>${esc(c)}</strong><br><span class='legend-detail'>${detail(c,m)}</span></div></div>`).join('')}</div>`,stackedLegend=cs=>`<div class='legend-list'>${cs.map((c,k)=>`<div class='legend-item'><span class='swatch-pair'><span class='swatch' style='background:${colors(k)}'></span><span class='swatch' style='background:hsl(${(k*137.508)%360} 75% 76%)'></span></span><div class='legend-copy'><strong>${esc(c)}</strong> <span class='legend-detail'>(dark = sent, light = received)</span><br><span class='legend-detail'>${detail(c,'volume')}</span></div></div>`).join('')}</div>`,mount=(title,draw)=>{const d=document.createElement('section');d.className='chart';d.innerHTML=`<h3>${title}</h3>`;const target=document.createElement('div');d.append(target);document.getElementById('charts').append(d);draw(target)};function lineChart(mode){const contacts=[...new Set(lines.map(r=>r.contact))],seasons=[...new Set(lines.map(r=>r.season))],W=Math.max(980,seasons.length*70),H=510,L=62,R=24,T=32,B=86,x=i=>L+i*(W-L-R)/Math.max(seasons.length-1,1),max=Math.max(...lines.map(r=>r.message_count),1),y=v=>mode==='rank'?T+(H-T-B)*(v-1)/Math.max(N-1,1):T+(H-T-B)*(1-v/max);let s=`<svg viewBox='0 0 ${W} ${H}'><text id='tip-${mode}' class='axis' x='${W-R}' y='18' text-anchor='end'>Hover a line</text>`;if(mode==='rank'){for(let v=1;v<=N;v++)s+=`<line class='grid' x1='${L}' x2='${W-R}' y1='${y(v)}' y2='${y(v)}'/><text class='axis' x='${L-8}' y='${y(v)+4}' text-anchor='end'>${v}</text>`}else{for(let i=0;i<=5;i++){let v=Math.round(max*i/5);s+=`<line class='grid' x1='${L}' x2='${W-R}' y1='${y(v)}' y2='${y(v)}'/><text class='axis' x='${L-8}' y='${y(v)+4}' text-anchor='end'>${v}</text>`}}seasons.forEach((q,i)=>s+=`<text class='axis' x='${x(i)}' y='${H-B+18}' text-anchor='end' transform='rotate(-45 ${x(i)} ${H-B+18})'>${q}</text>`);contacts.forEach((c,k)=>{const a=seasons.map(q=>lines.find(r=>r.season===q&&r.contact===c)),p=a.map((r,i)=>r?`${a[i-1]?'L':'M'}${x(i)},${y(mode==='rank'?r.rank:r.message_count)}`:'').join(' '),col=colors(k);if(p)s+=`<path stroke='transparent' stroke-width='18' fill='none' d='${p}' data-c='${k}'/><path class='line' stroke='${col}' d='${p}' data-c='${k}'/>`;a.forEach((r,i)=>r&&(s+=`<circle class='dot' fill='${col}' cx='${x(i)}' cy='${y(mode==='rank'?r.rank:r.message_count)}' r='4' data-c='${k}'/>`))});return s+'</svg>'+legend(contacts,mode)}mount('Top contacts by half-year - rank',d=>{d.innerHTML=lineChart('rank');wire(d,'rank')});mount('Top contacts by half-year - message volume',d=>{d.innerHTML=lineChart('volume');wire(d,'volume')});function wire(d,mode){const tip=d.querySelector('#tip-'+mode),cs=[...new Set(lines.map(r=>r.contact))];d.querySelectorAll('[data-c]').forEach(e=>{e.onmouseenter=()=>tip.textContent=cs[+e.dataset.c];e.onmouseleave=()=>tip.textContent='Hover a line'})}" & linefeed & Â
		"function stackedChart(){const contacts=[...new Set(stack.map(r=>r.contact))],seasons=[...new Set(stack.map(r=>r.season))],W=Math.max(980,seasons.length*70),H=510,L=62,R=24,T=32,B=86,x=i=>L+i*(W-L-R)/Math.max(seasons.length-1,1),tot=seasons.map(q=>stack.filter(r=>r.season===q).reduce((s,r)=>s+r.message_count,0)),max=Math.max(...tot,1),y=v=>T+(H-T-B)*(1-v/max);let s=`<svg viewBox='0 0 ${W} ${H}'>`;for(let i=0;i<=5;i++){let v=Math.round(max*i/5);s+=`<line class='grid' x1='${L}' x2='${W-R}' y1='${y(v)}' y2='${y(v)}'/><text class='axis' x='${L-8}' y='${y(v)+4}' text-anchor='end'>${v}</text>`}seasons.forEach((q,i)=>s+=`<text class='axis' x='${x(i)}' y='${H-B+18}' text-anchor='end' transform='rotate(-45 ${x(i)} ${H-B+18})'>${q}</text>`);let base=Array(seasons.length).fill(0);contacts.forEach((c,k)=>[['sent_count',colors(k)],['received_count',`hsl(${(k*137.508)%360} 75% 76%)`]].forEach(([f,col])=>{let vs=seasons.map(q=>stack.find(r=>r.season===q&&r.contact===c)[f]),top=vs.map((v,i)=>base[i]+v);s+=`<path class='area' fill='${col}' d='M ${top.map((v,i)=>x(i)+','+y(v)).join(' L ')} L ${base.map((v,i)=>x(i)+','+y(v)).reverse().join(' L ')} Z'/>`;base=top}));return s+'</svg>'+stackedLegend(contacts)}mount('Top contacts by half-year Ñ sent and received',d=>d.innerHTML=stackedChart());</script>"
end makeDashboardHTML
on sqlQuote(theText)
	set AppleScript's text item delimiters to "'"
	set textItems to text items of theText
	set AppleScript's text item delimiters to "''"
	set escapedText to textItems as text
	set AppleScript's text item delimiters to ""
	return "'" & escapedText & "'"
end sqlQuote
on writeText(theText, filePath)
	set fileRef to open for access (POSIX file filePath) with write permission
	try
		set eof of fileRef to 0
		write theText to fileRef as Çclass utf8È
		close access fileRef
	on error errMsg number errNum
		try
			close access fileRef
		end try
		error errMsg number errNum
	end try
end writeText