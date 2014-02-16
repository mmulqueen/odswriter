import re

class Formula(object):
    def __init__(self, s):
        self.tokens = lex_fragment(s)

    def __repr__(self):
        return "".join([self.__class__.__name__, "(", repr(self.tokens), ")"])

    def __str__(self):
        return "".join(["of:=",] + [str(token) for token in self.tokens])

class Token(object):
    def __repr__(self):
        return "".join([self.__class__.__name__, "(",
                        ",".join([repr(getattr(self, field)) for field in self.repr_fields]),
                        ")"])

class Function(Token):
    repr_fields = ["name", "arguments"]
    def __init__(self, name, argument_fragment):
        self.name = name
        self.arguments = lex_fragment(argument_fragment)

    def __str__(self):
        arg_str_parts = []
        last_index = len(self.arguments) - 1
        for i, arg in enumerate(self.arguments):
            prev_arg = None
            if i > 0:
                prev_arg = self.arguments[i-1]
            if not (isinstance(prev_arg, Operator)
                    or prev_arg is None
                    or isinstance(arg, Operator)):
                arg_str_parts.append(";")
            arg_str_parts.append(str(arg))

        return "".join([self.name, "(", "".join(arg_str_parts), ")"])

class Operator(Token):
    repr_fields = ["operator"]

    def __init__(self, operator):
        if operator == "==":
            self.operator = "="
        elif operator == "!=":
            self.operator = "<>"
        else:
            self.operator = operator

    def __str__(self):
        return self.operator


operator_tokens = ("!=", "<=", ">=", "<>", "==", "+", "-", "*", "/", "^", "%",
                   "&", "|", ">", "<", "=")
separator_tokens = (",",";")

def tokenize_fragment(s):
    if s.startswith("="):
        s = s[1:]
    special_char_pattern = re.escape("".join(operator_tokens + separator_tokens))
    pattern = re.compile("".join([r"\w*\(.+?\)|[\w\:]+|[",
                                  special_char_pattern,
                                  r"]+"]))
    return re.findall(pattern, s)

def lex_fragment(s):
    tokens = tokenize_fragment(s)
    lexed_tokens = []
    function_pattern = re.compile(r"[^\(\)]+")
    for token in tokens:
        token_parts = re.findall(function_pattern, token)
        if len(token_parts) > 1:
            lexed_tokens.append(Function(*token_parts))
        elif token in operator_tokens:
            lexed_tokens.append(Operator(token))
        elif token in separator_tokens:
            pass
        else:
            try:
                lexed_tokens.append(float(token))
            except ValueError:
                lexed_tokens.append(token)
    return lexed_tokens