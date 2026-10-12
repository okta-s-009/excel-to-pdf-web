# Keep JavaScriptInterface methods for AndroidBridge
-keepclassmembers class com.excel2pdf.studio.AndroidBridge {
    <methods>;
}

-keepattributes *Annotation*
-keepclassmembers class * {
    @android.webkit.JavascriptInterface <methods>;
}
