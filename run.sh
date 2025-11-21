#!/bin/bash
# Script to start spamcop_forwarder in a tmux session

SESSION="spamcop_forwarder"
SCRIPT_DIR="$HOME/code/spamcop_forwarder"
VENV_DIR="$SCRIPT_DIR/venv"
PYTHON_SCRIPT="$SCRIPT_DIR/spamcop_forwarder.py"

# Change to script directory
cd "$SCRIPT_DIR" || {
    echo "Error: Cannot change to directory $SCRIPT_DIR"
    exit 1
}

# Check if venv exists
if [ ! -d "$VENV_DIR" ]; then
    echo "Error: Virtual environment not found at $VENV_DIR"
    echo "Please create it first with: python3 -m venv venv"
    exit 1
fi

# Check if script exists
if [ ! -f "$PYTHON_SCRIPT" ]; then
    echo "Error: Script not found at $PYTHON_SCRIPT"
    exit 1
fi

# Check if tmux session already exists
if tmux has-session -t "$SESSION" 2>/dev/null; then
    echo "Session '$SESSION' already exists."
    echo "Use 'tmux attach -t $SESSION' to view it, or"
    echo "Use 'tmux kill-session -t $SESSION' to stop it first."
    exit 1
fi

# Start tmux session in detached mode
# The command activates venv and runs the script
tmux new-session -d -s "$SESSION" \
    "source $VENV_DIR/bin/activate && python3 $PYTHON_SCRIPT"

if [ $? -eq 0 ]; then
    echo "Successfully started spamcop_forwarder in tmux session: $SESSION"
    echo ""
    echo "To view the session: tmux attach -t $SESSION"
    echo "To detach: Press Ctrl+B, then D"
    echo "To stop: tmux kill-session -t $SESSION"
    echo ""
    echo "You can safely close this window - the session will continue running."
else
    echo "Error: Failed to start tmux session"
    exit 1
fi