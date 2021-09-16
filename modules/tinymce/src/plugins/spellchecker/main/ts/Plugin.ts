/**
 * Copyright (c) Tiny Technologies, Inc. All rights reserved.
 * Licensed under the LGPL or a commercial license.
 * For LGPL see License.txt in the project root for license information.
 * For commercial licenses see https://www.tiny.cloud/
 */

import { Cell } from '@ephox/katamari';
import PluginManager from 'tinymce/core/api/PluginManager';
import * as DetectProPlugin from './alien/DetectProPlugin';
import * as Api from './api/Api';
import * as Commands from './api/Commands';
import * as Settings from './api/Settings';
import * as Actions from './core/Actions';
import * as Buttons from './ui/Buttons';
import * as SuggestionsMenu from './ui/SuggestionsMenu';

type LastSuggestion = Actions.LastSuggestion;

export default () => {
  PluginManager.add('spellchecker', (editor, pluginUrl) => {
    if (DetectProPlugin.hasProPlugin(editor) === false) {
      const startedState = Cell(false);
      const currentLanguageState = Cell<string>(Settings.getLanguage(editor));
      const textMatcherState = Cell(null);
      const lastSuggestionsState = Cell<LastSuggestion>(null);
      const cache = Cell<any>(null);

      Buttons.register(editor, pluginUrl, startedState, textMatcherState, currentLanguageState, lastSuggestionsState, cache);
      SuggestionsMenu.setup(editor, pluginUrl, lastSuggestionsState, startedState, textMatcherState, currentLanguageState);
      Commands.register(editor, pluginUrl, startedState, textMatcherState, lastSuggestionsState, currentLanguageState, cache);

      Actions.setup(editor, pluginUrl, startedState, textMatcherState, lastSuggestionsState, currentLanguageState, cache);
      return Api.get(editor, startedState, lastSuggestionsState, textMatcherState, currentLanguageState, pluginUrl);
    }
  });
};
