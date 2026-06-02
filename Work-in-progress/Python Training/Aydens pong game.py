import turtle as t
import random
import math
import os

class PongGame:
    def __init__(self):
        # Game state
        self.player_a_score = 0
        self.player_b_score = 0
        self.game_paused = False
        self.winning_score = 10
        self.ball_speed = 3.0
        self.paddle_speed = 20
        
        # Key states for smooth movement
        self.keys_pressed = {
            'w': False, 's': False,
            'Up': False, 'Down': False
        }
        
        self.setup_window()
        self.create_game_objects()
        self.setup_controls()
        
    def setup_window(self):
        """Initialize the game window with better settings"""
        self.window = t.Screen()
        self.window.title("Enhanced Pong Game by Ayden Venter")
        self.window.bgcolor("#1a1a2e")  # Modern dark theme
        self.window.setup(width=1000, height=700)
        try:
            self.window.cv._rootwindow.resizable(False, False)  # Disable resizing
        except:
            pass  # Skip if not available
        self.window.tracer(0)
        
    def create_game_objects(self):
        """Create all game objects with improved styling"""
        # Left paddle (Player A - Ayden)
        self.left_paddle = t.Turtle()
        self.left_paddle.speed(0)
        self.left_paddle.shape("square")
        self.left_paddle.color("#00ff88")  # Neon green
        self.left_paddle.shapesize(stretch_wid=6, stretch_len=1.2)
        self.left_paddle.penup()
        self.left_paddle.goto(-450, 0)
        
        # Right paddle (Player B - Dadda)
        self.right_paddle = t.Turtle()
        self.right_paddle.speed(0)
        self.right_paddle.shape("square")
        self.right_paddle.color("#ff6b6b")  # Coral red
        self.right_paddle.shapesize(stretch_wid=6, stretch_len=1.2)
        self.right_paddle.penup()
        self.right_paddle.goto(450, 0)
        
        # Ball
        self.ball = t.Turtle()
        self.ball.speed(0)
        self.ball.shape("circle")
        self.ball.color("#ffffff")
        self.ball.penup()
        self.reset_ball()
        
        # Score display
        self.score_display = t.Turtle()
        self.score_display.speed(0)
        self.score_display.color("#ffffff")
        self.score_display.penup()
        self.score_display.hideturtle()
        self.score_display.goto(0, 300)
        
        # Game info display
        self.info_display = t.Turtle()
        self.info_display.speed(0)
        self.info_display.color("#888888")
        self.info_display.penup()
        self.info_display.hideturtle()
        self.info_display.goto(0, -320)
        
        # Center line
        self.draw_center_line()
        
        self.update_display()
        
    def draw_center_line(self):
        """Draw a dashed center line"""
        line_drawer = t.Turtle()
        line_drawer.speed(0)
        line_drawer.color("#444444")
        line_drawer.penup()
        line_drawer.hideturtle()
        line_drawer.goto(0, 320)
        line_drawer.setheading(270)
        
        # Draw dashed line
        for _ in range(32):
            line_drawer.pendown()
            line_drawer.forward(10)
            line_drawer.penup()
            line_drawer.forward(10)
    
    def play_sound(self, sound_type):
        """Play sound with fallback to system beep"""
        try:
            if sound_type == "paddle":
                os.system("afplay paddle.wav&")
            elif sound_type == "wall":
                os.system("afplay wallhit.wav&")
            elif sound_type == "score":
                os.system("afplay score.wav&")
        except:
            # Fallback - print sound notification (silent mode)
            pass
    
    def setup_controls(self):
        """Setup improved keyboard controls"""
        self.window.listen()
        
        # Key press events
        self.window.onkeypress(lambda: self.set_key('w', True), 'w')
        self.window.onkeypress(lambda: self.set_key('s', True), 's')
        self.window.onkeypress(lambda: self.set_key('Up', True), 'Up')
        self.window.onkeypress(lambda: self.set_key('Down', True), 'Down')
        
        # Key release events for smooth movement
        self.window.onkeyrelease(lambda: self.set_key('w', False), 'w')
        self.window.onkeyrelease(lambda: self.set_key('s', False), 's')
        self.window.onkeyrelease(lambda: self.set_key('Up', False), 'Up')
        self.window.onkeyrelease(lambda: self.set_key('Down', False), 'Down')
        
        # Game controls
        self.window.onkeypress(self.toggle_pause, 'space')
        self.window.onkeypress(self.reset_game, 'r')
        self.window.onkeypress(self.quit_game, 'Escape')
        
        # Alternative controls (W/S and Arrow keys)
        self.window.onkeypress(lambda: self.set_key('w', True), 'W')
        self.window.onkeypress(lambda: self.set_key('s', True), 'S')
    
    def set_key(self, key, state):
        """Set key state for smooth movement"""
        self.keys_pressed[key] = state
    
    def toggle_pause(self):
        """Toggle game pause state"""
        self.game_paused = not self.game_paused
        if self.game_paused:
            self.info_display.clear()
            self.info_display.write("PAUSED - Press SPACE to continue", 
                                  align="center", font=('Arial', 16, 'bold'))
        else:
            self.update_display()
    
    def reset_game(self):
        """Reset the entire game"""
        self.player_a_score = 0
        self.player_b_score = 0
        self.game_paused = False
        self.ball_speed = 3.0  # Reset ball speed
        self.reset_ball()
        self.update_display()
    
    def quit_game(self):
        """Clean exit"""
        try:
            self.window.bye()
        except:
            pass
    
    def reset_ball(self):
        """Reset ball to center with random direction"""
        self.ball.goto(0, 0)
        # Random angle between -45 and 45 degrees (or 135-225)
        angle = random.choice([
            random.uniform(-45, 45),
            random.uniform(135, 225)
        ])
        self.ball.dx = self.ball_speed * math.cos(math.radians(angle))
        self.ball.dy = self.ball_speed * math.sin(math.radians(angle))
    
    def update_display(self):
        """Update score and info displays"""
        # Score
        self.score_display.clear()
        self.score_display.write(f"Ayden: {self.player_a_score}    VS    Dadda: {self.player_b_score}", 
                               align="center", font=('Arial', 28, 'bold'))
        
        # Game info
        if not self.game_paused:
            self.info_display.clear()
            self.info_display.write(f"First to {self.winning_score} wins! | SPACE: Pause | R: Reset | ESC: Quit", 
                                  align="center", font=('Arial', 12, 'normal'))
    
    def handle_paddle_movement(self):
        """Handle smooth paddle movement"""
        # Left paddle (Ayden)
        if self.keys_pressed['w']:
            y = self.left_paddle.ycor() + self.paddle_speed
            if y < 280:  # Boundary check
                self.left_paddle.sety(y)
                
        if self.keys_pressed['s']:
            y = self.left_paddle.ycor() - self.paddle_speed
            if y > -280:  # Boundary check
                self.left_paddle.sety(y)
        
        # Right paddle (Dadda)
        if self.keys_pressed['Up']:
            y = self.right_paddle.ycor() + self.paddle_speed
            if y < 280:  # Boundary check
                self.right_paddle.sety(y)
                
        if self.keys_pressed['Down']:
            y = self.right_paddle.ycor() - self.paddle_speed
            if y > -280:  # Boundary check
                self.right_paddle.sety(y)
    
    def check_ball_collision(self):
        """Enhanced collision detection"""
        # Top and bottom walls
        if self.ball.ycor() > 320 or self.ball.ycor() < -320:
            self.ball.dy *= -1
            self.play_sound("wall")
            
            # Keep ball in bounds
            if self.ball.ycor() > 320:
                self.ball.sety(320)
            else:
                self.ball.sety(-320)
        
        # Right paddle collision
        if (420 < self.ball.xcor() < 460 and 
            self.right_paddle.ycor() - 60 < self.ball.ycor() < self.right_paddle.ycor() + 60):
            
            self.ball.setx(420)
            # Add spin based on where ball hits paddle
            hit_pos = (self.ball.ycor() - self.right_paddle.ycor()) / 60
            self.ball.dx = -abs(self.ball.dx)  # Ensure it goes left
            self.ball.dy += hit_pos * 2  # Add spin
            self.play_sound("paddle")
            
            # Increase speed slightly (max speed limit)
            if abs(self.ball.dx) < 8:
                self.ball.dx *= 1.05
                self.ball.dy *= 1.05
        
        # Left paddle collision  
        if (-460 < self.ball.xcor() < -420 and 
            self.left_paddle.ycor() - 60 < self.ball.ycor() < self.left_paddle.ycor() + 60):
            
            self.ball.setx(-420)
            # Add spin based on where ball hits paddle
            hit_pos = (self.ball.ycor() - self.left_paddle.ycor()) / 60
            self.ball.dx = abs(self.ball.dx)  # Ensure it goes right
            self.ball.dy += hit_pos * 2  # Add spin
            self.play_sound("paddle")
            
            # Increase speed slightly (max speed limit)
            if abs(self.ball.dx) < 8:
                self.ball.dx *= 1.05
                self.ball.dy *= 1.05
        
        # Scoring
        if self.ball.xcor() > 480:
            self.player_a_score += 1
            self.play_sound("score")
            self.reset_ball()
            self.update_display()
            self.check_winner()
            
        elif self.ball.xcor() < -480:
            self.player_b_score += 1
            self.play_sound("score")
            self.reset_ball()
            self.update_display()
            self.check_winner()
    
    def check_winner(self):
        """Check for game winner"""
        if self.player_a_score >= self.winning_score:
            self.game_paused = True
            self.info_display.clear()
            self.info_display.write("🎉 AYDEN WINS! 🎉 Press R to play again", 
                                  align="center", font=('Arial', 20, 'bold'))
        elif self.player_b_score >= self.winning_score:
            self.game_paused = True
            self.info_display.clear()
            self.info_display.write("🎉 DADDA WINS! 🎉 Press R to play again", 
                                  align="center", font=('Arial', 20, 'bold'))
    
    def run(self):
        """Main game loop with improved frame rate"""
        frame_count = 0
        try:
            while True:
                if not self.game_paused:
                    self.handle_paddle_movement()
                    
                    # Move ball
                    self.ball.setx(self.ball.xcor() + self.ball.dx)
                    self.ball.sety(self.ball.ycor() + self.ball.dy)
                    
                    # Check collisions
                    self.check_ball_collision()
                
                # Update screen less frequently for better performance
                frame_count += 1
                if frame_count % 1 == 0:  # Update every frame
                    try:
                        self.window.update()
                    except:
                        break  # Window was closed
                
        except (t.Terminator, TclError):
            pass  # Game window was closed
        except Exception as e:
            print(f"Game ended: {e}")

# Simple fallback version if the class doesn't work
def run_simple_pong():
    """Fallback simple pong game"""
    print("Running simplified version...")
    
    # Simple variables
    player_a_score = 0
    player_b_score = 0
    
    # Create window
    window = t.Screen()
    window.title("Pong Game by Ayden Venter")
    window.bgcolor("black")
    window.setup(width=800, height=600)
    window.tracer(0)
    
    # Left paddle
    left_paddle = t.Turtle()
    left_paddle.speed(0)
    left_paddle.shape("square")
    left_paddle.color("white")
    left_paddle.shapesize(stretch_wid=5, stretch_len=1)
    left_paddle.penup()
    left_paddle.goto(-350, 0)
    
    # Right paddle
    right_paddle = t.Turtle()
    right_paddle.speed(0)
    right_paddle.shape("square")
    right_paddle.color("white")
    right_paddle.shapesize(stretch_wid=5, stretch_len=1)
    right_paddle.penup()
    right_paddle.goto(350, 0)
    
    # Ball
    ball = t.Turtle()
    ball.speed(0)
    ball.shape("circle")
    ball.color("white")
    ball.penup()
    ball.goto(0, 0)
    ball.dx = 3
    ball.dy = 3
    
    # Score
    pen = t.Turtle()
    pen.speed(0)
    pen.color("white")
    pen.penup()
    pen.hideturtle()
    pen.goto(0, 260)
    pen.write("Ayden: 0    Dadda: 0", align="center", font=('Arial', 24, 'normal'))
    
    # Movement functions
    def left_paddle_up():
        y = left_paddle.ycor()
        if y < 250:
            left_paddle.sety(y + 20)
    
    def left_paddle_down():
        y = left_paddle.ycor()
        if y > -250:
            left_paddle.sety(y - 20)
    
    def right_paddle_up():
        y = right_paddle.ycor()
        if y < 250:
            right_paddle.sety(y + 20)
    
    def right_paddle_down():
        y = right_paddle.ycor()
        if y > -250:
            right_paddle.sety(y - 20)
    
    # Key bindings
    window.listen()
    window.onkeypress(left_paddle_up, 'w')
    window.onkeypress(left_paddle_down, 's')
    window.onkeypress(right_paddle_up, 'Up')
    window.onkeypress(right_paddle_down, 'Down')
    
    # Game loop
    while True:
        try:
            window.update()
            
            # Move ball
            ball.setx(ball.xcor() + ball.dx)
            ball.sety(ball.ycor() + ball.dy)
            
            # Ball collision with walls
            if ball.ycor() > 290:
                ball.sety(290)
                ball.dy *= -1
            
            if ball.ycor() < -290:
                ball.sety(-290)
                ball.dy *= -1
            
            # Ball goes off screen
            if ball.xcor() > 390:
                ball.goto(0, 0)
                ball.dx *= -1
                player_a_score += 1
                pen.clear()
                pen.write(f"Ayden: {player_a_score}    Dadda: {player_b_score}", 
                         align="center", font=('Arial', 24, 'normal'))
            
            if ball.xcor() < -390:
                ball.goto(0, 0)
                ball.dx *= -1
                player_b_score += 1
                pen.clear()
                pen.write(f"Ayden: {player_a_score}    Dadda: {player_b_score}", 
                         align="center", font=('Arial', 24, 'normal'))
            
            # Paddle collisions
            if (340 < ball.xcor() < 350 and 
                right_paddle.ycor() - 50 < ball.ycor() < right_paddle.ycor() + 50):
                ball.setx(340)
                ball.dx *= -1
            
            if (-350 < ball.xcor() < -340 and 
                left_paddle.ycor() - 50 < ball.ycor() < left_paddle.ycor() + 50):
                ball.setx(-340)
                ball.dx *= -1
                
        except:
            break

# Run the game
if __name__ == "__main__":
    print("Enhanced Pong Game")
    print("Controls:")
    print("Ayden (Left): W/S keys")  
    print("Dadda (Right): Up/Down arrow keys")
    print("SPACE: Pause/Unpause (Enhanced version)")
    print("R: Reset game (Enhanced version)")
    print("ESC: Quit (Enhanced version)")
    print("\nStarting game...")
    
    try:
        # Try the enhanced version first
        game = PongGame()
        game.run()
    except Exception as e:
        print(f"Enhanced version failed: {e}")
        print("Falling back to simple version...")
        try:
            run_simple_pong()
        except Exception as e2:
            print(f"Simple version also failed: {e2}")
            print("There may be an issue with your Python turtle installation.")