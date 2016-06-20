/// <reference path="wwwroot/tsd.d.ts" />

var gulp = require('gulp');
var del = require('del');
var tsb = require('gulp-tsb');
var browserSync = require('browser-sync').create();
var typescript = require('typescript');
var ts = require('gulp-typescript');
var sourcemaps = require('gulp-sourcemaps');

// TypeScript build using gulp-tsb
var tsb = tsb.create({
    target: 'es5',
    module: 'commonjs',
    sourceMap: true,
    declaration: false
},true);

gulp.task('tsb-build', function () {
    return gulp.src('wwwroot/**/*.ts')
        .pipe(tsb())
        .pipe(gulp.dest('wwwroot'));
});

// Typescript build using gulp-typescript
var projectConfig = {
    target: 'ES5',
    typescript: typescript

};
// var tsProject = ts.createProject('tsconfig.json');
gulp.task('build', function () {
    return gulp.src('wwwroot/**/*.ts', { base: 'wwwroot' })
        .pipe(sourcemaps.init())
        .pipe(ts(projectConfig))
        .pipe(sourcemaps.write('.', { includeContent: false, sourceRoot: '/wwwroot' }))
        .pipe(gulp.dest('wwwroot'));
});

gulp.task('clean', function (cb) {
    del('wwwroot/app/**/*.js');
    del('wwwroot/app/**/*.map');
});

gulp.task('reload', ['build'], function () {
    browserSync.reload();
});

gulp.task('serve', ['build'], function (done) {
    browserSync.init({
        // don't require online connectivity
        online: false,
        // don't open a browser
        open: false,
        port: 8080,
        server: {
            baseDir: ['wwwroot']
        }
    }, done);
});

gulp.task('watch', ['serve'], function (cb) {
    gulp.watch('wwwroot/**/*.ts', ['build', 'reload']);
});