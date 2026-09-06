function argbToHex(argb) {
    if (!argb || argb.length < 6) return '';
    // ExcelJS มักให้มาเป็น AARRGGBB หรือ RRGGBB
    return '#' + (argb.length === 8 ? argb.substring(2) : argb);
}

// ฟังก์ชันช่วยแกะรหัส cell (A1 -> {r, c})
function decodeAddr(addr) {
    const col = addr.match(/[A-Z]+/)[0];
    const row = parseInt(addr.match(/\d+/)[0]);
    let c = 0;
    for (let i = 0; i < col.length; i++) {
        c = c * 26 + col.charCodeAt(i) - 64;
    }
    return { r: row, c: c };
}


function getThemeColors(workbook) {
    const xml = workbook._themes?.theme1;
    if (!xml) return [];

    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const scheme = doc.getElementsByTagName("a:clrScheme")[0];
    if (!scheme) return [];

    const keys = [
        "lt1",
        "dk1",
        "lt2",
        "dk2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6"
    ];

    return keys.map(k => {
        const node = scheme.getElementsByTagName(`a:${k}`)[0];
        if (!node) return "FFFFFF";

        const srgb = node.getElementsByTagName("a:srgbClr")[0];
        if (srgb) return srgb.getAttribute("val");

        const sys = node.getElementsByTagName("a:sysClr")[0];
        if (sys) return sys.getAttribute("lastClr");

        return "FFFFFF";
    });
}

function getIndexedColor(index) {
    const palette = [
        "000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF",
        "000000","FFFFFF","000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00",
        "FF00FF","00FFFF","800000","008000","000080","808000","800080","008080",
        "C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080",
        "0066CC","CCCCFF","000080","FF00FF","FFFF00","00FFFF","800080","800000",
        "008080","0000FF","00CCFF","CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC",
        "CC99FF","FFCC99","3366FF","33CCCC","99CC00","FFCC00","FF9900","FF6600",
        "666699","969696","003366","339966","003300","333300","993300","993366",
        "000000","FFFFFF"
    ];

    return palette[index] || "000000";
}

function toARGB(color, themeColors) {
    if (!color) return null;

    // 1. ARGB เดิม
    if (color.argb) return color.argb;

    // 2. THEME
    if (color.theme !== undefined) {
        let hex = themeColors[color.theme];
        if (!hex) return null;

        let r = parseInt(hex.slice(0,2),16);
        let g = parseInt(hex.slice(2,4),16);
        let b = parseInt(hex.slice(4,6),16);

        const tint = color.tint || 0;

        const apply = (v) =>
            tint > 0 ? v + (255 - v) * tint : v * (1 + tint);

        r = Math.round(apply(r));
        g = Math.round(apply(g));
        b = Math.round(apply(b));

        return 'FF' +
            r.toString(16).padStart(2,'0') +
            g.toString(16).padStart(2,'0') +
            b.toString(16).padStart(2,'0');
    }

    // 3. INDEXED ⭐ เพิ่มตรงนี้
if (color.indexed !== undefined) {
    const hex = getIndexedColor(color.indexed);
    return "FF" + hex;
}

    return null;
}
