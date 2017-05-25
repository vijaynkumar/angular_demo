(function() {
  'use strict';

  angular
    .module('dumyProject')
    .run(runBlock);

  /** @ngInject */
  function runBlock($log) {

    $log.debug('runBlock end');
  }

})();
