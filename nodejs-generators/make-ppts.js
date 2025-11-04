
var XLSX = require('xlsx');
var PptxGenJS = require("pptxgenjs");

const befDefaults = {
    green : '94C601',
    darkGreen: '74A50F',
    black : '000000',
    white : 'FFFFFF',
    gray : '71685A',
    fontFace : 'Questrial'
};

function get_round_data_from_file(filename) {
    console.log("reading filename ", filename);
    const wb = XLSX.readFile(filename, { type: 'binary', WTF:1});

    console.log("Sheet names: ", wb.SheetNames);
    
    let rounds = {}
    
    Object.keys(wb.Sheets).map((ws_name) => {
	const sheet = wb.Sheets[ws_name];
	const matches = ws_name.match(/Round\s+(\d+)(\s|\+)?$/);
	if (matches){
	    const round_name = matches[1];
	    console.log(`Changing sheet name '${ws_name}' to round name '${round_name}'`);
	    rounds[round_name] = XLSX.utils.sheet_to_json(sheet);
	    console.log("Count of data rows:", rounds[round_name].length);
	    //console.log("first row of round data:", rounds[round_name].data[0]);
	} else {
	    console.log("skipping sheet named ",ws_name);
	}
    });
    return rounds;
}

function add_info_slide(pptx, text) {
    let slide = pptx.addNewSlide('BEF_MASTER_SLIDE');

    slide.addText(text,
		  { fontSize:36, fontFace:befDefaults.fontFace, color:befDefaults.black,
		    align:'c', x:'20%', y:3, w:8.5, w:'60%', h:2});
    return slide;
}

function add_end_round_slide(pptx, round_name) {
    let slide = pptx.addNewSlide('BEF_MASTER_SLIDE');

    slide.addText('This concludes\nRound ' + round_name, 
		  { fontSize:36, fontFace:befDefaults.fontFace, color:befDefaults.black,
		    align:'c', x:'20%', y:3, w:8.5, w:'60%', h:2});
    return slide;

}

function add_title_slide(pptx, opts) {
    let title_slide = pptx.addNewSlide('BEF_TITLE_SLIDE');

    title_slide.addText(opts.event_date,
			{x: "55%", y: "22%", w:"35%", h:"5%", color:befDefaults.white,
			 fontFace: befDefaults.fontFace, fontSize: 24, valign: 'top'});
    title_slide.addText('Brookline\nEducation\nFoundation',
			{x: "55%", y: "35%", w:"35%", h: "25%", align: 'l', valign:'top',
			 color:befDefaults.green, fontFace: befDefaults.fontFace, fontSize: 32});
    title_slide.addText('5th Grade\nSpelling Bee\n\nRound ' + opts.round,
			{x: "55%", y: "65%", w:"35%", h:"15%", bold: true, valign:'top',
			 color:befDefaults.black, fontFace: befDefaults.fontFace, fontSize: 22});
    return title_slide;
}

function add_slide_masters(pptx, opts) {

    pptx.defineSlideMaster( {
	title: 'BEF_TITLE_SLIDE',
	bkgd:   { fill: {type:'gradient', stops: [{ pos: 0, color:'C1F15E' },
						  { pos: 62, color:'90BA3F' },
						  { pos: 100, color:'7FA03E'}],
			 linearAngle: 90, linearScaled: false}},
	objects: [
	    { 'image':{x:0,y:0,w:"100%", h:"100%", path: 'bee-background.png',
		       sizing: {type: 'cover', x:0, y:0, w:"100%", h:"100%"}}},
	    { 'rect': {x:"50%", y:0, w:'41%', h:"90%", fill:befDefaults.white,
		       line:befDefaults.darkGreen, lineSize:1}},
	    { 'rect': {x:"51%", y:0, w:'39%', h:"30%",
		       fill:befDefaults.gray, line:'000000', lineSize:0}},
	    { 'rect': {x:"51%", y:"90%", w:'39%', h:2,
		       fill:befDefaults.green, line:'000000', lineSize:1}},
	    { 'image': { x:"70%", y:"60%", w:2.1, h:2.1, path:'bee-right.png' } }
	]
    });
//    { 'line': {x:'51%', y:'88%', w:'39%', h:1,line:befDefaults.darkGreen}},
    pptx.defineSlideMaster( {
	title:   'BEF_MASTER_SLIDE',
	bkgd:    {  fill: {type:'gradient',
			  stops: [{ pos: 0, color:'C1F15E' }, { pos: 62, color:'90BA3F' }, { pos: 100, color:'7FA03E'}],
			  linearAngle: 90, linearScaled: false}},
	margin:  [ 0.5, 0.25, 1.0, 0.25 ],
	objects: [
	    { 'image':{x:0,y:0,w:"100%",h:"100%", path: 'bee-background.png'}},
	    { 'rect': {x:"5%", y:"5%", w:'90%', h:"90%",
		       fill: befDefaults.white, line:'000000', lineSize:1}},
	    { 'image': { x:1, y:1.5, w:1.9, h:1.9, path:'bee-left.png' } },
	    { 'image': { x:6.9, y:1.5, w:1.9, h:1.9, path:'bee-right.png' } },
	    { 'rect': {x:"50%", y:0, w:'41%', h:"9%", fill:befDefaults.white, line:befDefaults.darkGreen, lineSize:1}},
	    { 'rect': {x:"51%", y:0, w:'39%', h:"8%", fill:befDefaults.gray, lineSize:0}},
	    { 'text':
	      {
		  text: '5th Grade Spelling Bee, Round ' + opts.round,
		  options: {x:"55%", y:0, w:'35%', h:0.6, align:'r', valign:'m',
			    color:befDefaults.white, fontSize:14, fontFace: befDefaults.fontFace }
	      }
	    },

	],
	//slideNumber: { x:0.6, y:7.0, color:befDefaults.black, fontFace:befDefaults.fontFace, fontSize:12 }
    });

}

function add_word_slide(pptx, word) {
    let slide = pptx.addNewSlide('BEF_MASTER_SLIDE');
    slide.addText( word,
		   { x: '20%', y: '40%', w:'60%', h:1,
		     color: befDefaults.black, align: 'c',
		     fontSize:48, fontFace:befDefaults.fontFace}
		 );
    return slide;
}

var seedrandom = require('seedrandom');

// shuffle array *inPlace* but returns array for chaining purposes
function shuffleArray(array, rng) {
    for (var i = array.length - 1; i > 0; i--) {
        var j = Math.floor(rng() * (i + 1));
        var temp = array[i];
        array[i] = array[j];
        array[j] = temp;
    }
    return array;
}

function make_round_pptx(round_records, opts)
{
    let pptx = new PptxGenJS();

    pptx.setAuthor(opts.author);
    pptx.setCompany(opts.company);

    var dateObj = new Date();
    var cur_month = dateObj.getMonth() + 1; //months from 1-12
    var cur_day = dateObj.getDate();
    var cur_year = dateObj.getFullYear();
    newdate = cur_year + "/" + cur_month + "/" + cur_day;

    pptx.setRevision(newdate);
    pptx.setSubject('5th Grade Spelling Bee Round ' + opts.round);
    pptx.setTitle('BEF 5th Grade Spelling Bee - ' + opts.year + ' - Round ' + opts.round);

    pptx.setLayout('LAYOUT_4x3');
    add_slide_masters(pptx, opts);

    console.log("creating new pptx");

    add_title_slide(pptx, opts);

    add_info_slide(pptx, 'Spellers at Work\n\n..quiet please...')
	.addText('Round ' + opts.round,
		 { fontSize:54, fontFace:befDefaults.fontFace,
		   color:befDefaults.green, align:'c',
		   x:'20%', y:1.25, w:'60%', h:1});

    // only find the words we care about
    word_records = round_records.filter(function(entry) {
	var keep = (entry.Word && entry[opts.year] != null && entry[opts.year] !== '');
	// console.log("filter is ", keep, " for entry ", entry);
	return keep;
    });
    var rng = seedrandom('hello.');
    rows = shuffleArray(word_records, rng);

    // loop over words we are going to output
    var len = word_records.length
    for(let ii = 0; ii < len; ii++) {
	const word = word_records[ii].Word;
	//console.log("adding slide for word:", word);
	add_word_slide(pptx, word);
	if (ii < len - 1) {
	    add_info_slide(pptx, 'Spellers at Work\n\n..quiet please...');
	}
    }
    add_info_slide(pptx, 'This concludes\nRound ' + opts.round);

    return pptx;
}

function main(){
    const year = process.argv[2];
    const xlsx_filename = process.argv[3];
    const output_dir = process.argv[4];
    let rounds = get_round_data_from_file(xlsx_filename);
    //let year = '19Fall';
    Object.keys(rounds).forEach(function(round_name) {
	let round_data = rounds[round_name];
	const opts = { round: round_name,
		       year: year,
		       event_date: 'Nov 17, 2024',
		       author: 'Perry A Stoll <perry@pstoll.com>',
		       company: 'Brookline Education Foundation'
		     };
	let pptx = make_round_pptx(round_data, opts);
	const outfilename = `${output_dir}/bee-slides-${year}-round${round_name}.pptx`;
	console.log(`saving ${round_name} pptx to ${outfilename}`);
	pptx.save(outfilename);
    });
}

main();
