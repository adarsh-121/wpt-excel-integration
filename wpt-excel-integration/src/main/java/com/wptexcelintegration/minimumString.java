import java.util.Stack;

public class MinimumStringLength {

	public static int findMinimumLength(String seq) {
		Stack<Character> stack = new Stack<>();

		for (char c : seq.toCharArray()) {

			if (!stack.isEmpty() && stack.peek == 'A' && c == 'B') {
				stack.pop();
			} else {
				stack.push(c);
			}
		}

		return stack.size();
	}

	public static void main(String[] args) {

		String seq = "ABABBABBAB";
		int minimumLength = fingMinimumLength(seq);
		System.out.println("Minimum Length of the remaining string " + minimumLength);
	}
}