/**
 * Copyright (c) Tiny Technologies, Inc. All rights reserved.
 * Licensed under the LGPL or a commercial license.
 * For LGPL see License.txt in the project root for license information.
 * For commercial licenses see https://www.tiny.cloud/
 */

import { Cell, Obj } from '@ephox/katamari';
import Editor from 'tinymce/core/api/Editor';
import Tools from 'tinymce/core/api/util/Tools';
import URI from 'tinymce/core/api/util/URI';
import XHR from 'tinymce/core/api/util/XHR';
import * as Events from '../api/Events';
import * as Settings from '../api/Settings';
import { DomTextMatcher } from './DomTextMatcher';

export interface Data {
  words: Record<string, string[]>;
  dictionary?: any;
}

const getTextMatcher = (editor, textMatcherState) => {
  if (!textMatcherState.get()) {
    const textMatcher = DomTextMatcher(editor.getBody(), editor);
    textMatcherState.set(textMatcher);
  }

  return textMatcherState.get();
};

const defaultSpellcheckCallback = (editor: Editor, pluginUrl: string, currentLanguageState: Cell<string>) => {
  return (method: string, text: string, doneCallback: Function, errorCallback: Function) => {
    const data = { method, lang: currentLanguageState.get() };
    let postData = '';

    data[method === 'addToDictionary' ? 'word' : 'text'] = text;

    Tools.each(data, (value, key) => {
      if (postData) {
        postData += '&';
      }

      postData += key + '=' + encodeURIComponent(value);
    });
    XHR.send({
      url: new URI(pluginUrl).toAbsolute(Settings.getRpcUrl(editor)),
      type: 'post',
      content_type: 'application/x-www-form-urlencoded',
      data: postData,
      success: (result) => {
        const parseResult = JSON.parse(result);

        if (!parseResult) {
          const message = editor.translate(`Server response wasn't proper JSON.`);
          errorCallback(message);
        } else if (parseResult.error) {
          errorCallback(parseResult.error);
        } else {
          doneCallback(parseResult);
        }
      },
      error: () => {
        const message = editor.translate('The spelling service was not found: (') +
          Settings.getRpcUrl(editor) +
          editor.translate(')');
        errorCallback(message);
      }
    });
  };
};

const sendRpcCall = (editor: Editor, pluginUrl: string, currentLanguageState: Cell<string>, name: string, data: string, words: string[], successCallback: Function, errorCallback?: Function) => {
  const userSpellcheckCallback = Settings.getSpellcheckerCallback(editor);
  if (userSpellcheckCallback) {
    userSpellcheckCallback.call(editor.plugins.spellchecker, name, words, successCallback, errorCallback);
  } else {
    defaultSpellcheckCallback(editor, pluginUrl, currentLanguageState).call(editor.plugins.spellchecker, name, data, successCallback, errorCallback);
  }
};

const spellcheck = (editor: Editor, pluginUrl: string, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, lastSuggestionsState: Cell<LastSuggestion>, currentLanguageState: Cell<string>, cache: Cell<any>, keepStarted = false) => {
  const lruCache = getCache(cache);

  if (finish(editor, startedState, textMatcherState, keepStarted)) {
    lruCache.clear();
    return;
  }

  const words = new Set<string>();
  const result: Record<string, string[]> = {};
  const minWordLength = Settings.getSpellcheckerMinWordLength(editor);
  const text = getTextMatcher(editor, textMatcherState).text;
  (text.match(Settings.getSpellcheckerWordcharPattern(editor)) || []).forEach((word) => {
    if (word.length >= minWordLength) {
      const suggestions = lruCache.get(word);
      if (suggestions) {
        result[word] = suggestions;
      } else if (lruCache.contains(word)) {
        // null est la valeur pour un mot ok
        //
      } else if (!words.has(word)) {
        words.add(word);
      }
    }
  });

  if (!words.size) {
    markErrors(editor, startedState, textMatcherState, lastSuggestionsState, { words: result, dictionary: [] });
  } else {
    const errorCallback = (message: string) => {
      editor.notificationManager.open({ text: message, type: 'error' });
      editor.setProgressState(false);
      finish(editor, startedState, textMatcherState);
    };

    const successCallback = (data: Data) => {
      words.forEach((word) => {
        const suggestions = data.words[word];
        if (suggestions) {
          lruCache.put(word, suggestions);
          result[word] = suggestions;
        } else {
          // null pour un mot ok
          //
          lruCache.put(word, null);
        }
      });

      markErrors(editor, startedState, textMatcherState, lastSuggestionsState, { words: result, dictionary: [] });
    };

    editor.setProgressState(true);
    sendRpcCall(editor, pluginUrl, currentLanguageState, 'spellcheck', text, Array.from(words), successCallback, errorCallback);
    editor.focus();
  }
};

const checkIfFinished = (editor: Editor, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>) => {
  if (!editor.dom.select('span.mce-spellchecker-word').length) {
    finish(editor, startedState, textMatcherState);
  }
};

const addToDictionary = (editor: Editor, pluginUrl: string, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, currentLanguageState: Cell<string>, word: string, spans: Element[]) => {
  editor.setProgressState(true);

  sendRpcCall(editor, pluginUrl, currentLanguageState, 'addToDictionary', word, [ word ], () => {
    editor.setProgressState(false);
    editor.dom.remove(spans, true);
    checkIfFinished(editor, startedState, textMatcherState);
  }, (message) => {
    editor.notificationManager.open({ text: message, type: 'error' });
    editor.setProgressState(false);
  });
};

const ignoreWord = (editor: Editor, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, word: string, spans: Element[], all?: boolean) => {
  editor.selection.collapse();

  if (all) {
    Tools.each(editor.dom.select('span.mce-spellchecker-word'), (span) => {
      if (span.getAttribute('data-mce-word') === word) {
        editor.dom.remove(span, true);
      }
    });
  } else {
    editor.dom.remove(spans, true);
  }

  checkIfFinished(editor, startedState, textMatcherState);
};

const finish = (editor: Editor, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, keepStarted = false) => {
  const bookmark = editor.selection.getBookmark();
  getTextMatcher(editor, textMatcherState).reset();
  editor.selection.moveToBookmark(bookmark);

  textMatcherState.set(null);

  if (startedState.get() && !keepStarted) {
    startedState.set(false);
    Events.fireSpellcheckEnd(editor);
    return true;
  }
};

const getElmIndex = (elm: HTMLElement) => {
  const value = elm.getAttribute('data-mce-index');

  if (typeof value === 'number') {
    return '' + value;
  }

  return value;
};

const findSpansByIndex = (editor: Editor, index: string): HTMLSpanElement[] => {
  const spans: HTMLSpanElement[] = [];

  const nodes = Tools.toArray(editor.getBody().getElementsByTagName('span'));
  if (nodes.length) {
    for (let i = 0; i < nodes.length; i++) {
      const nodeIndex = getElmIndex(nodes[i]);

      if (nodeIndex === null || !nodeIndex.length) {
        continue;
      }

      if (nodeIndex === index.toString()) {
        spans.push(nodes[i]);
      }
    }
  }

  return spans;
};

export interface LastSuggestion {
  suggestions: Record<string, string[]>;
  hasDictionarySupport: boolean;
}

const markErrors = (editor: Editor, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, lastSuggestionsState: Cell<LastSuggestion>, data: Data) => {
  const hasDictionarySupport = !!data.dictionary;
  const suggestions = data.words;

  editor.setProgressState(false);

  const empty = Obj.isEmpty(suggestions);
  const dynamic = Settings.getSpellcheckerDynamic(editor);
  if (empty && !dynamic) {
    const message = editor.translate('No misspellings found.');
    editor.notificationManager.open({ text: message, type: 'info' });
    startedState.set(false);
    return;
  }

  lastSuggestionsState.set({
    suggestions,
    hasDictionarySupport
  });

  if (!empty) {
    const bookmark = editor.selection.getBookmark();

    getTextMatcher(editor, textMatcherState).find(Settings.getSpellcheckerWordcharPattern(editor)).filter((match) => {
      return !!suggestions[match.text];
    }).wrap((match) => {
      return editor.dom.create('span', {
        'class': 'mce-spellchecker-word',
        'aria-invalid': 'spelling',
        'data-mce-bogus': 1,
        'data-mce-word': match.text
      });
    });

    editor.selection.moveToBookmark(bookmark);
  }

  if (!startedState.get()) {
    startedState.set(true);
    Events.fireSpellcheckStart(editor);
  }
};

const setup = (editor: Editor, pluginUrl: string, startedState: Cell<boolean>, textMatcherState: Cell<DomTextMatcher>, lastSuggestionsState: Cell<LastSuggestion>, currentLanguageState: Cell<string>, cache: Cell<any>): void => {
  const cacheSize = Settings.getSpellcheckerCacheSize(editor);
  cache.set(new LruCache<string[]>(cacheSize));

  /* eslint-disable no-console */
  console.debug('spellchecker cache size set to ' + cacheSize);
  console.debug('spellchecker min word length set to ' + Settings.getSpellcheckerMinWordLength(editor));
  /* eslint-enable */

  if (!Settings.getSpellcheckerDynamic(editor)) {
    return;
  }

  const dynamicDelay = Settings.getSpellcheckerDynamicDelay(editor);
  const dynamicSpaceDelay = Settings.getSpellcheckerDynamicSpaceDelay(editor);

  // eslint-disable-next-line no-console
  console.debug('spellchecker configured as dynamic with delay set to ' + dynamicDelay + ' ms and space delay set to ' + dynamicSpaceDelay + ' ms');

  let timeout = 0;

  const inputListener = (e) => {
    let delay = 0;
    switch (e.inputType) {
      case 'insertText':
      case 'insertReplacementText':
      case 'insertFromYank':
      case 'insertFromDrop':
      case 'insertFromPaste':
      case 'insertFromPasteAsQuotation':
      case 'insertTranspose':
      case 'insertCompositionText':
      case 'insertLink':
      case 'deleteWordBackward':
      case 'deleteWordForward':
      case 'deleteSoftLineBackward':
      case 'deleteSoftLineForward':
      case 'deleteEntireSoftLine':
      case 'deleteHardLineBackward':
      case 'deleteHardLineForward':
      case 'deleteByDrag':
      case 'deleteByCut':
      case 'deleteContent':
      case 'deleteContentBackward':
      case 'deleteContentForward':
      case 'historyUndo':
      case 'historyRedo':
        delay = (e.data && (e.data.indexOf(' ') >= 0)) ? 500 : 1500;
        break;

      case 'insertLineBreak':
      case 'insertParagraph':
      case 'insertOrderedList':
      case 'insertUnorderedList':
      case 'insertHorizontalRule':
      case 'formatBold':
      case 'formatItalic':
      case 'formatUnderline':
      case 'formatStrikeThrough':
      case 'formatSuperscript':
      case 'formatSubscript':
      case 'formatJustifyFull':
      case 'formatJustifyCenter':
      case 'formatJustifyRight':
      case 'formatJustifyLeft':
      case 'formatIndent':
      case 'formatOutdent':
      case 'formatRemove':
      case 'formatSetBlockTextDirection':
      case 'formatSetInlineTextDirection':
      case 'formatBackColor':
      case 'formatFontColor':
      case 'formatFontName':
        break;
    }

    if (delay > 0) {
      clearTimeout(timeout);
      timeout = setTimeout(() => {
        if (startedState.get()) {
          spellcheck(editor, pluginUrl, startedState, textMatcherState, lastSuggestionsState, currentLanguageState, cache, true);
        }
      }, delay);
    }
  };

  editor.on('spellcheckstart spellcheckend', (event) => {
    if (event.type === 'spellcheckstart') {
      editor.on('input', inputListener);
    } else {
      editor.off('input', inputListener);
    }
  });
};

const getCache = (cache: Cell<any>): LruCache<string[]> => cache.get();

class LruCache<T> {
  private values: Map<string, T> = new Map<string, T>();
  private maxEntries: number;

  public constructor(maxEntries: number) {
    this.maxEntries = maxEntries;
  }

  public get(key: string): T {
    const hasKey = this.values.has(key);
    let entry: T;
    if (hasKey) {
      // peek the entry, re-insert for LRU strategy
      //
      entry = this.values.get(key);
      this.values.delete(key);
      this.values.set(key, entry);
    }
    return entry;
  }

  public put(key: string, value: T): void {
    if (this.maxEntries > 0) {
      if (this.values.size >= this.maxEntries) {
        // least-recently used cache eviction strategy
        //
        const keyToDelete = this.values.keys().next().value;
        this.values.delete(keyToDelete);
      }

      this.values.set(key, value);
    }
  }

  public contains(key: string): boolean {
    return this.values.has(key);
  }

  public clear() {
    this.values.clear();
  }
}

export {
  setup,
  spellcheck,
  checkIfFinished,
  addToDictionary,
  ignoreWord,
  findSpansByIndex,
  getElmIndex,
  markErrors
};
