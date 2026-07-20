import streamlit as st 
import clipboard_component as cc


if "str_text" not in st.session_state:
    st.session_state["str_text"] = ""

if "str_textAppend" not in st.session_state:
    st.session_state["str_textAppend"] = ""

if "str_option1" not in st.session_state:
    st.session_state["str_option1"] = ""

if "str_option2" not in st.session_state:
    st.session_state["str_option2"] = ""

if "str_option3" not in st.session_state:
    st.session_state["str_option3"] = ""

if "str_option4" not in st.session_state:
    st.session_state["str_option4"] = ""

if "letterAns" not in st.session_state:
    st.session_state["letterAns"] = ""

if "indexAns" not in st.session_state:
    st.session_state["indexAns"] = 0


# Question
text = cc.paste_component("Paste Question")

if text:
    st.session_state["str_text"] = text

text = st.text_input("Enter Question: ", key= "str_text")

st.write(st.session_state["str_text"])

# Question Append
textAppend = cc.paste_component("Paste append question")

if textAppend:
    st.session_state["str_textAppend"] = textAppend

textAppend = st.text_input("Enter question to append:", key= "str_textAppend")

st.write(st.session_state["str_textAppend"])

# Letter A
option1 = cc.paste_component("Paste A")

if option1:
    st.session_state["str_option1"] = option1

option1 = st.text_input("Enter Letter A: ", key= "str_option1")

st.write(st.session_state["str_option1"])

# Letter B
option2 = cc.paste_component("Paste B")

if option2:
    st.session_state["str_option2"] = option2

option2 = st.text_input("Enter Letter B: ", key= "str_option2")

st.write(st.session_state["str_option2"])

# Letter C
option3 = cc.paste_component("Paste C")

if option3:
    st.session_state["str_option3"] = option3

option3 = st.text_input("Enter Letter C: ", key= "str_option3")

st.write(st.session_state["str_option3"])

# Letter D
option4 = cc.paste_component("Paste D")

if option4:
    st.session_state["str_option4"] = option4

option4 = st.text_input("Enter Letter D: ", key= "str_option4")

st.write(st.session_state["str_option4"])

# Correct Answer
col1, col2, col3, col4 = st.columns(4, gap= None)

with col1:
    if st.button("A",use_container_width=True):
        st.session_state["letterAns"] = "A"
        st.session_state["indexAns"] = 0

with col2:
    if st.button("B", use_container_width=True):
        st.session_state["letterAns"] = "B"
        st.session_state["indexAns"] = 1

with col3:
    if st.button("C", use_container_width=True):
        st.session_state["letterAns"] = "C"
        st.session_state["indexAns"] = 2

with col4:
    if st.button("D", use_container_width=True):
        st.session_state["letterAns"] = "D"
        st.session_state["indexAns"] = 3

st.write(st.session_state["letterAns"])

# Return Value
formatted = """{ "text": """ + repr(st.session_state["str_text"] + " " + st.session_state["str_textAppend"]) + ", " + """ "options": [""" + repr("A. " + st.session_state["str_option1"] ) + ", " + repr("B. " + st.session_state["str_option2"]) + ", " + repr("C. " + st.session_state["str_option3"]) + ", " + repr("D. " + st.session_state["str_option4"]) +"]," + """ "correct": """ + str(st.session_state["indexAns"]) + "}, "

st.code(formatted, language=None)
