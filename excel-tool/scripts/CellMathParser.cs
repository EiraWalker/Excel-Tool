using System;
using System.Collections.Generic;
using System.Text;

public sealed class CellMathRun
{
    public string Text;
    public string Kind;
    public CellMathRun(string text, string kind) { Text = text; Kind = kind; }
}

// A deliberately limited inline notation parser, not an OfficeMath renderer.
public sealed class CellMathParser
{
    private readonly string source;
    private int position;
    private readonly List<CellMathRun> runs = new List<CellMathRun>();
    private static readonly Dictionary<string, string> Symbols = new Dictionary<string, string> {
        {"alpha", "α"}, {"beta", "β"}, {"gamma", "γ"}, {"delta", "δ"},
        {"epsilon", "ε"}, {"theta", "θ"}, {"lambda", "λ"}, {"mu", "μ"},
        {"pi", "π"}, {"rho", "ρ"}, {"sigma", "σ"}, {"tau", "τ"},
        {"phi", "φ"}, {"omega", "ω"}, {"Delta", "Δ"}, {"Theta", "Θ"},
        {"Lambda", "Λ"}, {"Pi", "Π"}, {"Sigma", "Σ"}, {"Omega", "Ω"},
        {"times", "×"}, {"cdot", "·"}, {"le", "≤"}, {"leq", "≤"},
        {"ge", "≥"}, {"geq", "≥"}, {"ne", "≠"}, {"neq", "≠"},
        {"in", "∈"}, {"land", "∧"}, {"lor", "∨"}, {"pm", "±"},
        {"infty", "∞"}, {"prime", "′"}
    };

    private CellMathParser(string value) { source = value; }
    private FormatException Error(string message) {
        return new FormatException(message + " at character " + (position + 1));
    }
    private void Add(string text, string kind) {
        if (text.Length == 0) return;
        if (runs.Count > 0 && runs[runs.Count - 1].Kind == kind)
            runs[runs.Count - 1].Text += text;
        else runs.Add(new CellMathRun(text, kind));
    }
    private string Atom() {
        if (position >= source.Length) throw Error("Missing script or command argument");
        char c = source[position++];
        if (c == '{') {
            var text = new StringBuilder();
            while (position < source.Length && source[position] != '}') text.Append(Atom());
            if (position == source.Length) throw Error("Unclosed group");
            position++;
            if (text.Length == 0) throw Error("Empty group");
            return text.ToString();
        }
        if (c == '}' || c == '^' || c == '_') throw Error("Nested scripts or stray braces are unsupported");
        if (c == '$') throw Error("Provide math without outer delimiters");
        if (c == '\\') {
            if (position == source.Length) throw Error("Incomplete command");
            if (!Char.IsLetter(source[position])) {
                char escaped = source[position++];
                if (escaped == ',' || escaped == ' ') return " ";
                if (escaped == '{' || escaped == '}' || escaped == '_' || escaped == '%' || escaped == '#') return escaped.ToString();
                throw Error("Unsupported escape");
            }
            int start = position;
            while (position < source.Length && Char.IsLetter(source[position])) position++;
            string command = source.Substring(start, position - start);
            string symbol;
            if (Symbols.TryGetValue(command, out symbol)) return symbol;
            if (command == "mathrm" || command == "text") {
                if (position >= source.Length || source[position] != '{') throw Error("Command requires a group");
                return Atom();
            }
            if (command == "quad") return "  ";
            throw Error("Unsupported cell-math command: \\" + command);
        }
        if (Char.IsSurrogate(c) || Char.IsControl(c)) throw Error("Use BMP characters and separate text segments for line breaks");
        return c == '-' ? "−" : c.ToString();
    }
    public static CellMathRun[] Parse(string value) {
        if (String.IsNullOrEmpty(value)) throw new FormatException("Math expression is empty");
        if (value.Length > 20000) throw new FormatException("Math expression is too long");
        var parser = new CellMathParser(value);
        bool hasBase = false, hasSub = false, hasSup = false;
        while (parser.position < value.Length) {
            char c = value[parser.position];
            if (c == '_' || c == '^') {
                if (!hasBase) throw parser.Error("Script requires a preceding base");
                if (hasSub || hasSup) throw parser.Error("Combined scripts on one base are unsupported");
                parser.position++;
                parser.Add(parser.Atom(), c == '_' ? "sub" : "sup");
                if (c == '_') hasSub = true; else hasSup = true;
            } else {
                string atom = parser.Atom();
                parser.Add(atom, "math");
                hasBase = !String.IsNullOrWhiteSpace(atom);
                hasSub = false; hasSup = false;
            }
        }
        return parser.runs.ToArray();
    }
}
