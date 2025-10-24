const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

// for reading spreadsheet of data
var XLSX = require('xlsx');

const fs = require("fs");
const path = require("path");

var expressions= require('angular-expressions');

/**
 * Randomize array element order in-place.
 * Using Durstenfeld shuffle algorithm.
 * https://stackoverflow.com/questions/2450954/how-to-randomize-shuffle-a-javascript-array
 */

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


function bold_sentence_docx(example, word) {
    var new_example = example;

    var word_offset = example.toLowerCase().indexOf(word.toLowerCase());
    // keep case of word in the example sentence.
    if ( word_offset !== -1) {
        // we found the word, so split it in pieces
        // make word found in example bold.
        exampleParts = [ example.slice(0, word_offset), 
                         example.slice(word_offset, word_offset + word.length),
                         example.slice(word_offset + word.length)];

	var font_str = '<w:sz w:val="40"/><w:szCs w:val="40"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>';
	var new_example = '<w:p>' +
	    '<w:pPr><w:rPr><w:i/>' + font_str + '</w:rPr></w:pPr>' + 
	    '<w:r><w:rPr><w:i/>' + font_str + '</w:rPr><w:t xml:space="preserve">' + exampleParts[0] + '</w:t></w:r>' +
	    '<w:r><w:rPr><w:i/><w:b/>' + font_str + '</w:rPr><w:t xml:space="preserve">' + exampleParts[1] + ' </w:t></w:r>' + 
	    '<w:r><w:rPr><w:i/>' + font_str + '</w:rPr><w:t>' + exampleParts[2] + '</w:t></w:r>' +
	    '</w:p>';
    } else {
	console.log(`Could not find ${word} in sentence: ${example}`)
    }
    return new_example;
}

// define your filter functions here, for example, to be able to write {clientname | lower}
//function regexEscape(str) {
//    return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
//} 

// var angularParser = function(tag) {
//     return {
//         get: tag === '.' ? function(s){ return s;} : function(s) {
//             return expressions.compile(tag.replace(/’/g, "'"))(s);
//         }
//     };
// }
//doc.setOptions({parser:angularParser});

//Load the JSON data from the file

function filter_round_by_year(records, yr) {
    var my_year_records = records.filter(function (el) {
	var val = el[yr];
	var is_empty = (val==null || val.match("^\s*$"));
	// console.log(`filtering year record ${yr} with value '${val}' - is_empty=${is_empty}`);
	return !is_empty;
    });
    return my_year_records;
}

function load_round_data_from_csv(filepath) {
    var bee_text = fs.readFileSync(path.resolve(__dirname, filepath), 'utf8');
    var bee_data = JSON.parse(bee_text);
    return bee_data;
}

// handle spaces at start or end of entries as necessary
function clean_word_entry(entry) {
    entry.Word= entry.Word.trim();
    entry.Pronunciation = entry.Pronunciation.trim();
    return entry;
}

function prepare_round_data(round_records) {

    console.log("Found ", round_records.length, " input json records");
    console.log("first row entry: ", round_records[0]);

    var valid_records = round_records.filter(function (el) {
	return (el.Word && el.Pronunciation && el.Definition && el.Sentence);
    });
    valid_records.map(clean_word_entry);

    console.log("Have ", valid_records.length, " filtered records");
    return valid_records;
}


function get_bee_data_xlsx(xls_file) {

    console.log("reading filename ", filename);
    var wb = XLSX.readFile(filename, { type: 'binary'});

    console.log("Raw Sheet Names: ", wb.SheetNames);

    var rounds = {}

    Object.keys(wb.Sheets).map((ws_name) => {
	const sheet = wb.Sheets[ws_name];
	const matches = ws_name.match(/Round\s+(\d+)(\s|\+)?$/);
	if (matches){
	    const round_name = matches[1];
	    console.log(`Changing sheet name '${ws_name}' to round name '${round_name}'`);
	    var records = XLSX.utils.sheet_to_json(sheet);
	    rounds[round_name] = prepare_round_data(records);
	    console.log("Count of data rows:", rounds[round_name].length);
	    console.log("first row of round data:", rounds[round_name][0]);
	} else {
	    console.log("skipping sheet named ",ws_name);
	}
    });
    console.log("done reading records. keys are ", Object.keys(rounds));
    return rounds;

}
function get_bee_data() {
    var rounds = {};
    for(var round = 1; round <= 6; round++) {
	var filename =  path.resolve('./output/','bee-words-round' + round + '.json');
	console.log("Getting data for round ", round, " at filename ", filename);
	var json_records = load_round_data(filename);
	rounds[round] = prepare_round_data(json_records);
	console.log("Loaded data for round ", round);
    }
    return rounds;
}


function generate_round_bee_doc_buf( cur_year, cur_round, rows, template_content) {
    // instatiate new doc template writer

    const zip = new PizZip(template_content);

    const doc = new Docxtemplater(zip, {
	paragraphLoop: true,
	linebreaks: true,
    });

    doc.setData( { "entries": rows, "withHeader": true, Year: cur_year, Round: cur_round} );
    try {
	// render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
	doc.render()
    }
    catch (error) {
	var e = {
	    message: error.message,
	    name: error.name,
	    stack: error.stack,
	    properties: error.properties,
	}
	console.log(JSON.stringify({error: e}));
	// The error thrown here contains additional information when logged with JSON.stringify
	// (it contains a property object).
	throw error;

    }
    var buf = doc.getZip().generate({
	type: 'nodebuffer',
	compression: "DEFLATE",
    });
    return buf;

}

function generate_year_bee_outputs( template, rounds, cur_year, output_dir, output_ext) {
    console.log("rounds.length=", Object.keys(rounds).length);

    for(var cur_round in rounds) {

	// filter makes a copy, so we can mutate returned values
	var rows = filter_round_by_year(rounds[cur_round], cur_year);
	var rng = seedrandom('hello.');
	rows = shuffleArray(rows, rng);

	console.log("Have ", rows.length, " records for year ", cur_year);

	// docx specific fix: adjust raw word xml data to manually bold word in sentence.
	for(var ii = 0; ii < rows.length; ii++) {
	    rows[ii].Sentence = bold_sentence_docx(rows[ii].Sentence, rows[ii].Word);
	}

	var buf = generate_round_bee_doc_buf(cur_year, cur_round, rows, template);
	// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
	var outfilename = `${output_dir}/bee-words-${cur_year}-round${cur_round}.${output_ext}`;
	console.log("Saving output to ", outfilename);
	fs.writeFileSync(path.resolve(__dirname, outfilename), buf);

    }

}

var cur_year = process.argv[2]; //'2019Fall';
var filename = process.argv[3];
var template_name = process.argv[4];
var output_dir = process.argv[5] // "outputs"
var output_ext = template_name.split('.').pop();

// Load the docx file as binary content once
const template_content = fs.readFileSync(
    path.resolve(__dirname, template_name),
    "binary"
);

var bee_data = get_bee_data_xlsx(filename);
//var bee_data = get_bee_data();
generate_year_bee_outputs(template_content, bee_data, cur_year, output_dir, output_ext);



