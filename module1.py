#import tkinter

#from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
## Implement the default Matplotlib key bindings.
#from matplotlib.backend_bases import key_press_handler
#from matplotlib.figure import Figure

#import numpy as np


#root = tkinter.Tk()
#root.wm_title("Embedding in Tk")

#fig = Figure(figsize=(5, 4), dpi=100)
#t = np.arange(0, 3, .01)
#fig.add_subplot(111).plot(t, 2 * np.sin(2 * np.pi * t))

#canvas = FigureCanvasTkAgg(fig, master=root)  # A tk.DrawingArea.
#canvas.draw()
#canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

#toolbar = NavigationToolbar2Tk(canvas, root)
#toolbar.update()
#canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)


#def on_key_press(event):
#    print("you pressed {}".format(event.key))
#    key_press_handler(event, canvas, toolbar)


#canvas.mpl_connect("key_press_event", on_key_press)


#def _quit():
#    root.quit()     # stops mainloop
#    root.destroy()  # this is necessary on Windows to prevent
#                    # Fatal Python Error: PyEval_RestoreThread: NULL tstate


#button = tkinter.Button(master=root, text="Quit", command=_quit)
#button.pack(side=tkinter.BOTTOM)

#tkinter.mainloop()
## If you put root.destroy() here, it will cause an error if the window is
## closed with the window manager.

# library
#import matplotlib.pyplot as plt

## Data to plot
#labels = 'Chanfrein_1', 'Coup_Denture_1', 'Chanfrein_2', 'Autres'
#sizes = [215, 205, 140, 25]
#colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
#explode = (0, 0, 0, 0)  # explode 1st slice

## Plot
#plt.pie(sizes, explode=explode, labels=labels, colors=colors,autopct='%1.1f%%', shadow=True, startangle=140)

#plt.axis('equal')
#plt.show()

import matplotlib as mpl
import numpy as np
import sys
if sys.version_info[0] < 3:
    import Tkinter as tk
else:
    import tkinter as tk
import matplotlib.backends.tkagg as tkagg
from matplotlib.backends.backend_agg import FigureCanvasAgg


def draw_figure(canvas, figure, loc=(0, 0)):
    """ Draw a matplotlib figure onto a Tk canvas

    loc: location of top-left corner of figure on canvas in pixels.
    Inspired by matplotlib source: lib/matplotlib/backends/backend_tkagg.py
    """
    figure_canvas_agg = FigureCanvasAgg(figure)
    figure_canvas_agg.draw()
    figure_x, figure_y, figure_w, figure_h = figure.bbox.bounds
    figure_w, figure_h = int(figure_w), int(figure_h)
    photo = tk.PhotoImage(master=canvas, width=figure_w, height=figure_h)

    # Position: convert from top-left anchor to center anchor
    canvas.create_image(loc[0] + figure_w/2, loc[1] + figure_h/2, image=photo)

    # Unfortunately, there's no accessor for the pointer to the native renderer
    tkagg.blit(photo, figure_canvas_agg.get_renderer()._renderer, colormode=2)

    # Return a handle which contains a reference to the photo object
    # which must be kept live or else the picture disappears
    return photo

# Create a canvas
w, h = 300, 200
window = tk.Tk()
window.title("A figure in a canvas")
canvas = tk.Canvas(window, width=w, height=h)
canvas.pack()

# Generate some example data
X = np.linspace(0, 2 * np.pi, 50)
Y = np.sin(X)

# Create the figure we desire to add to an existing canvas
fig = mpl.Figure(figsize=(2, 1))
ax = fig.add_axes([0, 0, 1, 1])
ax.plot(X, Y)

# Keep this handle alive, or else figure will disappear
fig_x, fig_y = 100, 100
fig_photo = draw_figure(canvas, fig, loc=(fig_x, fig_y))
fig_w, fig_h = fig_photo.width(), fig_photo.height()

# Add more elements to the canvas, potentially on top of the figure
canvas.create_line(200, 50, fig_x + fig_w / 2, fig_y + fig_h / 2)
canvas.create_text(200, 50, text="Zero-crossing", anchor="s")

# Let Tk take over
tk.mainloop()