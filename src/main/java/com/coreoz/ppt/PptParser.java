package com.coreoz.ppt;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import lombok.AllArgsConstructor;

class PptParser {
	
	/*this is returning the text between $/ / and there maybe an additional argument */
	static Optional<PptVariable> parse(String text) {
		if(text.startsWith("$/") && text.endsWith("/")) {
			int indexStartParameter = text.indexOf(':');
			if(indexStartParameter < 0) {
				return Optional.of(PptVariable.of(text.substring(2, text.length() - 1), null));
			}
			return Optional.of(PptVariable.of(
				text.substring(2, indexStartParameter),
				text.substring(indexStartParameter + 1, text.length() - 1)
			));
		}
		return Optional.empty();
	}
	
	
	
	static void replaceTextVariable(XSLFTextParagraph paragraph, PptMapper mapper) {
		int indexOfStartVariable = -1;
		List<XSLFTextRun> textPartsVariable = null;
		StringBuilder variableName = null;
		StringBuilder variableArgument = null;
		State currentState = State.INITIAL;

		/* An XSLFTextParagraph can have one or more XSLFTextRun
		 * we are going through the text in each TextRun  */
		for(XSLFTextRun textPart : paragraph.getTextRuns()) {
			char[] textPartRaw = textPart.getRawText().toCharArray();
			int indexOfChar = 0;  //this is always getting incremented for all textruns
			
			//now we are going through each char of the text from the XSLFTextRun
			//we could be in different states depending on the current char we are scanning
			
			if(currentState.inVariable) {
				textPartsVariable.add(textPart);
			}

			for(char c : textPartRaw) {
				State nextState = processCharAndGetNextState(currentState, c);

				switch (nextState) {
				case INITIAL:
					/*we were going through maybe a variable but we are back to initial state.. so parser has not found a expected match.  */
					if(currentState != State.INITIAL) {
						indexOfStartVariable = -1;
						textPartsVariable = null;
						variableName = null;
						variableArgument = null;
					}

					break;
				case MAY_BE_VARIABLE:
					indexOfStartVariable = indexOfChar;
					textPartsVariable = new ArrayList<>();
					textPartsVariable.add(textPart);

					break;
				case START_VARIABLE:
					variableName = new StringBuilder();

					break;
				case VARIABLE:
					variableName.append(c);

					break;
				case START_ARGUMENT:
					variableArgument = new StringBuilder();

					break;
				case ARGUMENT:
					variableArgument.append(c);

					break;
				case END_VARIABLE:
					indexOfChar = replaceVariable(
						indexOfStartVariable,
						indexOfChar, //this will be last char.. we need to replace from indexOfStartVariable to indexOfChar
						mapper.textMapping(
							variableName.toString(),
							variableArgument == null ? null : variableArgument.toString()
						),   /*this is returning the value which we want to use i.e this value should be used to replace the variable */
						textPartsVariable
					);
					break;
				}

				indexOfChar++;
				currentState = nextState;
			}
		}
	}
	
	

	/**
	 *
	 * @param indexOfStartVariable The index of the first char of the variable in the first TextRun
	 * @param indexOfEndVariable The index of the last char of the variable in the last TextRun
	 * @param replacedText The value to replace the variable
	 * @param textParts The text parts in which the variable name should be replaced by its value
	 * @return The index of the character in the last text part to continue to search for variable
	 */
	private static int replaceVariable(int indexOfStartVariable, int indexOfEndVariable,
			Optional<String> replacedText, List<XSLFTextRun> textParts) {
		
		if(!replacedText.isPresent()) {
			return indexOfEndVariable;
		}
		
		//going through all the textParts
		// alll these textParts are of the same XSLFTextParagraph
		for (int i = 0; i < textParts.size(); i++) {
			XSLFTextRun textPart = textParts.get(i);
			
			if(i == 0) {
				String partContent = textPart.getRawText();
				StringBuilder textPartReplaced = new StringBuilder(partContent.substring(0, indexOfStartVariable));
				textPartReplaced.append(replacedText.get());
				
				if(textParts.size() == 1) {
					textPartReplaced.append(partContent.substring(indexOfEndVariable + 1)); //the remaining text
				}
				textPart.setText(textPartReplaced.toString());
	
				if(textParts.size() == 1) {
					return replacedText.get().length() - 1;
				}
			} else if(i < (textParts.size() - 1)) {
				textPart.setText("");
			} else {
				textPart.setText(textPart.getRawText().substring(indexOfEndVariable + 1));
				return -1;
			}
		}

		throw new RuntimeException("Parsing issue");
	}

	
	
	private static State processCharAndGetNextState(State before, char c) {
		switch (before) {
		case END_VARIABLE:
			
			
		/*If the next char starts with $ then it Maybe a variable */	
		case INITIAL:
			if(c == '$') {
				return State.MAY_BE_VARIABLE;
			}
			break;
			
		/*If the next char after $ is / then it is start of variable  */	
		case MAY_BE_VARIABLE:
			if(c == '/') {
				return State.START_VARIABLE;
			}
			break;
		
			/*if after / we dont find another / then we are going through the variable. */
		case START_VARIABLE:
			if(c != '/') {
				return State.VARIABLE;
			}
			break;
		/*if we find another / then its end of variable, but if we find a : then its start of an argument passed */	
		case VARIABLE:
			if(c == '/') {
				return State.END_VARIABLE;
			}
			if(c == ':') {
				return State.START_ARGUMENT;
			}
			return State.VARIABLE;
		case START_ARGUMENT: /*we fallthrough the case */
		case ARGUMENT:
			if(c == '/') {
				return State.END_VARIABLE;
			}
			return State.ARGUMENT;
		}

		return State.INITIAL;
	}
	
	

	@AllArgsConstructor
	private static enum State {
		INITIAL(false),    /*when we have just started scanning the text */
		MAY_BE_VARIABLE(true),
		START_VARIABLE(true),
		VARIABLE(true),
		START_ARGUMENT(true),
		ARGUMENT(true),
		END_VARIABLE(false)
		;

		private boolean inVariable;
	}

}
