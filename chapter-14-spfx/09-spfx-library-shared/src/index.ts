// Public surface of the shared SPFx Library.
// Consumed by other solutions via `import { ... } from 'spfx-library-shared';`
export { SharedLibraryLibrary } from './libraries/sharedLibrary/SharedLibraryLibrary';
export { classifyHttpStatus, HttpStatusCategory } from './libraries/sharedLibrary/utils/httpStatus';
export { truncate, formatRelativeDate } from './libraries/sharedLibrary/utils/format';