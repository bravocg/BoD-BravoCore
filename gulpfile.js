var gulp = require("gulp");
var concat = require("gulp-concat");
var rename = require("gulp-rename");
var uglify = require("gulp-uglify");

gulp.task("dist", function () {
    return gulp.src("./bravo.js")
        .pipe(gulp.dest("dist"))
        .pipe(rename("bravo.min.js"))
        .pipe(uglify())
        .pipe(gulp.dest("dist"));
});

gulp.task("default", ["dist"]);