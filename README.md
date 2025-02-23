# IO_STEPFileReader  
## parsing, tokenizing, reading and writing files in step format aka ISO 10303-21  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/IO_STEPFileReader?style=plastic)](https://github.com/OlimilO1402/IO_STEPFileReader/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/IO_STEPFileReader?style=plastic)](https://github.com/OlimilO1402/IO_STEPFileReader/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/IO_STEPFileReader/total.svg)](https://github.com/OlimilO1402/IO_STEPFileReader/releases/download/v2025.02.23/StepReader_v2025.02.23.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)


Project started in january 2025. The idea started about 6 years ago.   
The step format is a human readable aka "ansi/ascii-text-based" file format.
It is for serializing and deserializing all kinds of data coming from object oriented hierarchies, and is widely used in the global industry.
e.g. CAD-data or building information modelling by using industry foundation classes (ifc).
The fact that it is human readable text based, makes it necessary for parsing and tokenizing in the first place to be able to read the objects in the file.

The specification is not freely available as also noted in the Wikipedia article. 
You can pay several hundreds of Euro for the ISO standard just by ending up with a pile of heavy readable papers to digging through.
Or you *are* the *developer* You are the *master*, the *hero* of your own code, so switch on your brain, have a deep look inside the file, make advantage from it's human readablity and learn how it's done even from one file alone.
At first glance it looks very easy, and you do not need the ISO at all. A second thought on it, it is similar to writing a parser for a programming language.  
You need some knowledge in the theory of state machines resp finite state automatons.  

```vba
Public Function NextToken() As StepToken
	'
End Function
```
Some Links:  
===========
[Wikipedia ISO 10303-21](https://en.wikipedia.org/wiki/ISO_10303-21) 
[Wikipedia Industry Foundation Classes](https://en.wikipedia.org/wiki/Industry_Foundation_Classes)
  
[Youtube channel Intermation](https://www.youtube.com/@Intermation)
[Intermation: Introduction to State Machines: Designing a Simple Traffic Signal](https://www.youtube.com/watch?v=gv5fQrD8XUo)
[Intermation: Introduction to Finite State Machine Theory](https://www.youtube.com/watch?v=05zp5vbYXSQ)
[Intermation: Converting Non-Deterministic Finite Automata to Deterministic Finite Automata](https://www.youtube.com/watch?v=wJk82yQhxms)
[Intermation: Finite State Automata - From Theory to Code](https://www.youtube.com/watch?v=IhUqXgVl6jo)
  
![Picture Image](Resources/Picture.png "Picture Image")
