package business;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Validator {

    /**
     * Takes a string and returns it if string is a valid email address. If not a valid email address, a null string will be returned.
     * @param s String: email address to be validated
     * @return String that either contains a valid email address or is null
     */
    public static String validateEmail(String s) {
        String validEmail = null;
        Pattern emailPattern = Pattern.compile("([A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4})", Pattern.CASE_INSENSITIVE);
        Matcher emailMatcher = emailPattern.matcher(s);
        if (emailMatcher.find()) {
            validEmail = emailMatcher.group(1);
        }
        return validEmail;
    }
}
