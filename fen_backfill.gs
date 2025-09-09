// FEN Backfill (resumable) for "Openings Normalized"
// - Loads chess.js dynamically (non-ESM builds)
// - Converts PGN -> final FEN
// - Writes FEN and split fields (board/active/castle/ep/halfmove/fullmove + ranks r8..r1)
// - Processes in batches and resumes via ScriptProperties

var __CHESS_CTOR__ = null;

function FEN_ensureChessLoaded_() {
  if (__CHESS_CTOR__) return;

  var urls = [
    'https://cdnjs.cloudflare.com/ajax/libs/chess.js/0.13.4/chess.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/chess.js/0.10.3/chess.min.js',
    'https://cdn.jsdelivr.net/npm/chess.js@0.13.4/chess.min.js'
  ];

  for (var i = 0; i < urls.length; i++) {
    try {
      var res = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() !== 200) continue;

      var code = res.getContentText();
      if (/\bexport\s/.test(code)) continue; // skip ESM builds

      // Strategy 1: sandboxed factory that returns the constructor
      try {
        var factory = new Function(
          'GLOBAL',
          '"use strict";' +
          'var window=GLOBAL,self=GLOBAL,global=GLOBAL;' +
          'var module={exports:{}};var exports=module.exports;' +
          code + ';' +
          'return GLOBAL.Chess || GLOBAL.chess || ' +
          '(module && module.exports && (module.exports.Chess || module.exports.default || module.exports));'
        );
        var result = factory({});
        if (typeof result === 'function') { __CHESS_CTOR__ = result; return; }
        if (result && typeof result.Chess === 'function') { __CHESS_CTOR__ = result.Chess; return; }
      } catch (e) {
        // fall through
      }

      // Strategy 2: plain eval, then read global
      try {
        eval(code);
        if (typeof Chess === 'function') { __CHESS_CTOR__ = Chess; return; }
        if (typeof chess === 'function') { __CHESS_CTOR__ = chess; return; }
      } catch (e2) {
        // try next url
      }
    } catch (e3) {
      // try next url
    }
  }

  throw new Error('Could not load a usable chess.js build from fallback URLs.');
}

function FEN_newChess_() {
  FEN_ensureChessLoaded_();
  return new __CHESS_CTOR__();
}

function FEN_sanitizePgn_(pgn) {
  pgn = String(pgn).replace(/\r/g, '').replace(/^\uFEFF/, '');
  pgn = pgn.replace(/^\s*\[[^\]]*\]\s*$/mg, '');       // headers
  pgn = pgn.replace(/\{[^}]*\}/g, '');                    // {...} comments
  for (var pass = 0; pass < 3; pass++) {                    // (...) variations
    pgn = pgn.replace(/\([^()]*\)/g, '');
  }
  pgn = pgn.replace(/\u2026/g, '...');                     // ellipsis
  pgn = pgn.replace(/\$\d+/g, '');                        // NAGs
  pgn = pgn.replace(/\b(\d+)\s+(?=[^.\s])/g, '$1. ');   // "1 e4" -> "1. e4"
  pgn = pgn.replace(/([KQRBNOa-h][^()\s]*)[!?]+/g, '$1');  // strip annotations
  pgn = pgn.replace(/([KQRBNOa-h0-9=+#-]+)[,;]+/g, '$1');   // strip punctuation
  pgn = pgn.replace(/\s+/g, ' ').trim();                   // collapse ws
  pgn = pgn.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/i, ''); // strip result
  return pgn;
}

function FEN_manualReplayFens_(pgn) {
  var tokens = pgn.split(/\s+/);
  var cleaned = [];
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (/^\d+\.+$/.test(t) || /^\d+\.\.\.$/.test(t) || /^\d+$/.test(t)) continue;     // move numbers
    if (/^(1-0|0-1|1\/2-1\/2|\*)$/i.test(t)) continue;                                    // results
    t = t.replace(/^\d+\.(\.\.)?/, '');                                                   // embedded numbers
    t = t.replace(/[;,]+$/g, '');
    if (t) cleaned.push(t);
  }
  var game = FEN_newChess_();
  var fens = [];
  for (var j = 0; j < cleaned.length; j++) {
    var san = cleaned[j];
    var moved = game.move(san, { sloppy: true });
    if (!moved) {
      var retry = san.replace(/[!?]+$/g, '');
      if (retry !== san) moved = game.move(retry, { sloppy: true });
    }
    if (!moved) break;
    fens.push(game.fen());
  }
  return fens;
}

function FEN_pgnToFens(pgn) {
  if (Array.isArray(pgn)) {
    var parts = [];
    for (var r = 0; r < pgn.length; r++) {
      var row = pgn[r];
      for (var c = 0; c < row.length; c++) if (row[c] != null) parts.push(String(row[c]));
    }
    pgn = parts.join('\n');
  } else {
    pgn = String(pgn);
  }

  var cleaned = FEN_sanitizePgn_(pgn);

  try {
    var game = FEN_newChess_();
    if (game.load_pgn(cleaned, { sloppy: true })) {
      var moves = game.history({ verbose: true });
      var replay = FEN_newChess_();
      var fens = [];
      for (var i = 0; i < moves.length; i++) {
        replay.move(moves[i]);
        fens.push(replay.fen());
      }
      return fens;
    }
  } catch (e) {
    // fall through
  }

  return FEN_manualReplayFens_(cleaned);
}

function FEN_pgnToFinalFen_(pgn) {
  try {
    var fens = FEN_pgnToFens(pgn);
    return fens.length ? fens[fens.length - 1] : FEN_newChess_().fen();
  } catch (e) {
    try {
      var game = FEN_newChess_();
      if (game.load_pgn(String(pgn), { sloppy: true })) return game.fen();
    } catch (e2) {}
    return '';
  }
}

function FEN_splitFen_(fen) {
  fen = String(fen || '').trim();
  if (!fen) {
    return {
      board: '', active: '', castle: '', ep: '', halfmove: '', fullmove: '',
      ranks: ['', '', '', '', '', '', '', '']
    };
  }
  var parts = fen.split(/\s+/);
  var board = parts[0] || '';
  var active = parts[1] || '';
  var castle = parts[2] || '';
  var ep = parts[3] || '';
  var halfmove = parts[4] || '';
  var fullmove = parts[5] || '';
  var ranks = board.split('/');
  while (ranks.length < 8) ranks.push('');
  if (ranks.length > 8) ranks = ranks.slice(0, 8);
  return { board: board, active: active, castle: castle, ep: ep, halfmove: halfmove, fullmove: fullmove, ranks: ranks };
}

var FEN_BF = { propRow: 'FEN_BF_NEXT_ROW', batchRows: 400 };

function FEN_getOrCreateTargetSheetEnsuringFenHeaders_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Openings Normalized');
  if (!sheet) {
    sheet = ss.insertSheet('Openings Normalized');
  }
  if (sheet.getLastRow() === 0) {
    // If empty, attempt to set at least the known base headers from existing importer if present
    var base = ['Family','Variation','Subvariation 1','Subvariation 2','Subvariation 3','ECO','Name','PGN','SourceFile','DataLine','Key'];
    sheet.getRange(1, 1, 1, base.length).setValues([base]);
    sheet.setFrozenRows(1);
  }

  // Ensure FEN-related columns are present; append any that are missing to the far right
  var required = [
    'FEN','FEN_board','FEN_active','FEN_castle','FEN_ep','FEN_halfmove','FEN_fullmove',
    'FEN_r8','FEN_r7','FEN_r6','FEN_r5','FEN_r4','FEN_r3','FEN_r2','FEN_r1'
  ];

  var lastCol = sheet.getLastColumn() || 1;
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (currentHeaders.indexOf(required[i]) === -1) missing.push(required[i]);
  }
  if (missing.length) {
    sheet.insertColumnsAfter(lastCol, missing.length);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function FEN_backfillResume() {
  var sheet = FEN_getOrCreateTargetSheetEnsuringFenHeaders_();
  if (!sheet) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colPGN = headers.indexOf('PGN') + 1;
  var colFEN = headers.indexOf('FEN') + 1;
  var colBoard = headers.indexOf('FEN_board') + 1;
  var colActive = headers.indexOf('FEN_active') + 1;
  var colCastle = headers.indexOf('FEN_castle') + 1;
  var colEp = headers.indexOf('FEN_ep') + 1;
  var colH = headers.indexOf('FEN_halfmove') + 1;
  var colF = headers.indexOf('FEN_fullmove') + 1;
  var colR8 = headers.indexOf('FEN_r8') + 1;

  if (colPGN <= 0 || colFEN <= 0 || colBoard <= 0 || colActive <= 0 || colCastle <= 0 || colEp <= 0 || colH <= 0 || colF <= 0 || colR8 <= 0) return;

  var props = PropertiesService.getScriptProperties();
  var startRow = parseInt(props.getProperty(FEN_BF.propRow) || '2', 10);
  var lastRow = sheet.getLastRow();
  if (startRow > lastRow) return;

  var endRow = Math.min(lastRow, startRow + FEN_BF.batchRows - 1);

  var numRows = endRow - startRow + 1;
  var pgns = sheet.getRange(startRow, colPGN, numRows, 1).getValues();
  var fens = sheet.getRange(startRow, colFEN, numRows, 1).getValues();

  var updFEN = new Array(numRows);
  var updBoard = new Array(numRows);
  var updActive = new Array(numRows);
  var updCastle = new Array(numRows);
  var updEp = new Array(numRows);
  var updH = new Array(numRows);
  var updF = new Array(numRows);
  var updRanks = new Array(numRows);

  for (var i = 0; i < numRows; i++) {
    var pgn = (pgns[i][0] || '').toString();
    var fenCell = (fens[i][0] || '').toString();

    var fen = fenCell;
    if (!fen && pgn) {
      try { fen = FEN_pgnToFinalFen_(pgn); } catch (e) { fen = ''; }
    }
    updFEN[i] = [fen];

    var sp = FEN_splitFen_(fen);
    updBoard[i] = [sp.board];
    updActive[i] = [sp.active];
    updCastle[i] = [sp.castle];
    updEp[i] = [sp.ep];
    updH[i] = [sp.halfmove];
    updF[i] = [sp.fullmove];
    updRanks[i] = [sp.ranks[0], sp.ranks[1], sp.ranks[2], sp.ranks[3], sp.ranks[4], sp.ranks[5], sp.ranks[6], sp.ranks[7]];
  }

  sheet.getRange(startRow, colFEN, numRows, 1).setValues(updFEN);
  sheet.getRange(startRow, colBoard, numRows, 1).setValues(updBoard);
  sheet.getRange(startRow, colActive, numRows, 1).setValues(updActive);
  sheet.getRange(startRow, colCastle, numRows, 1).setValues(updCastle);
  sheet.getRange(startRow, colEp, numRows, 1).setValues(updEp);
  sheet.getRange(startRow, colH, numRows, 1).setValues(updH);
  sheet.getRange(startRow, colF, numRows, 1).setValues(updF);
  sheet.getRange(startRow, colR8, numRows, 8).setValues(updRanks);

  PropertiesService.getScriptProperties().setProperty(FEN_BF.propRow, String(endRow + 1));
}

function FEN_resetBackfillProgress() {
  PropertiesService.getScriptProperties().deleteProperty(FEN_BF.propRow);
}

// Optional trigger helpers specific to FEN backfill
function FEN_createMinuteTrigger() {
  ScriptApp.newTrigger('FEN_backfillResume').timeBased().everyMinutes(1).create();
}

function FEN_deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === 'FEN_backfillResume') {
      ScriptApp.deleteTrigger(t);
    }
  }
}

