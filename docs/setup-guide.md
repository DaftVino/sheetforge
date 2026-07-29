# SheetForge — Setup Guide

## What this is

SheetForge adds a set of buttons to your own Excel file. You pick the ones you
want on a setup page. You do not write any code.

This guide takes about ten minutes. You do not need to know anything about
Excel macros.

---

## The one thing to know first

Read this before you start. It is the part that confuses people.

> To change your settings, click `Finalize Setup Tab`. Then click
> `Create Setup Tab`. That makes the new questions show up. Answer them. Then
> click `Finalize Setup Tab` again to save them.

So it is a loop. Pick what you want. Save it. Get asked more. Answer. Save
again. You will click `Finalize Setup Tab` twice in the steps below. That is on
purpose. It is not a mistake in this guide.

---

## The steps

Do these in order. After each step you will see something new. That is how you
know it worked.

### 1. Get the file

Go to the [Releases page](../../releases/latest). Click the file whose name ends
in `.bas`. Click **Download**.

You now have one file. It is the whole toolkit.

### 2. Open your Excel file

Open the workbook you want the buttons in. Any workbook will do.

Save it as a `.xlsm` file first. Click **File**, then **Save As**, then pick
**Excel Macro-Enabled Workbook**. A plain `.xlsx` file cannot hold macros. It
will drop them when you close it, and Excel will not warn you.

### 3. Open the code window

Press `Alt` and `F11` at the same time. A new window opens. It looks nothing
like Excel. That is fine. You will only be here for a moment.

### 4. Bring the file in

Click **File** in the top left. Click **Import File...**

Find the file you got in step 1. Click it. Click **Open**.

Nothing seems to happen. That is normal.

### 5. Go back to Excel

Press `Alt` and `Q` at the same time. You are back in your workbook.

### 6. Run the first macro

Press `Alt` and `F8` at the same time. A small box opens with a list of names.

The list has three names in it:

- `Core_BuildSetup` — makes the setup page. **This is the one you want now.**
- `Core_FinalizeSetup` — saves your answers and builds your buttons. You will
  do this with a button later, so you do not need it from this list.
- `Core_CompleteSetup` — tidies up at the very end. That is step 12.

Click `Core_BuildSetup`. Click **Run**.

### 7. Fill in the yellow boxes

Two new tabs appear at the bottom of your screen. One is called `Setup`. The
other is called `Setup Menu`. You will use both.

Go to the `Setup` tab. Some cells on it are yellow. Yellow means "fill this in".

Go down the page and fill in the ones you need. You can leave the rest blank
for now.

### 8. Pick your macros

Scroll down to the list of macros. Each one has an `Availability` cell next to
it. Click that cell and a small arrow appears. Click the arrow to get a list of
three choices:

- **Button** — you get a button for it.
- **Chains only** — no button, but you can use it in a chain.
- **Off** — you do not get it.

Every macro starts as **Off**. Pick **Button** for each one you want.

A line of text shows up above the list as soon as you pick your first one. It
says:

> Actions are switched on. Click 'Finalize Setup Tab' and then 'Create Setup
> Tab' to reveal their option fields.

That is the loop from the top of this page. You are about to do it.

### 9. Save what you picked

Go to the `Setup Menu` tab. Click `Finalize Setup Tab`.

A box says **Configuration activated**. Click **OK**. You land on a new tab
called `Macros Home`.

This step does two jobs. It saves your answers, and it builds your buttons. So
if you see that box, it worked.

### 10. Answer the new questions

Go back to the `Setup Menu` tab. Click `Create Setup Tab`.

New yellow boxes show up on the `Setup` tab. They are the questions your chosen
macros need answered. A macro that saves files will ask where to put them.

Answer them. Then go to `Setup Menu` and click `Finalize Setup Tab` again.

You get the same **Configuration activated** box, and you land on `Macros Home`
again. Now your buttons know your answers.

### 11. Use your buttons

Your buttons are on the `Macros 1` tab. If you picked a lot of them, there will
be more tabs: `Macros 2`, and so on.

There is also a tab called `Macros Home`. It lists every macro, what it does,
and which tab its button is on. It lists all of them, even the ones you left
**Off**. Go there when you forget what something does.

That is it. You are set up. You can stop here.

### 12. When you are done for good

Only do this when your settings are final. It is a one-way step.

Go to the `Setup Menu` tab. Click `Complete Macro Setup`.

This takes away the `Setup`, `Setup Menu` and `Macros Home` tabs. Your buttons
stay, and they still work.

It asks you first. Read the box. Press **Enter** and the answer is no, so
nothing gets deleted. You have to click **Yes** on purpose.

If you need the setup page back later, press `Alt` and `F8`, then run
`Core_BuildSetup` again. Your answers will still be there.

---

## Three small asks

**Ask for a macro you want.** Open the [Issues page](../../issues) and click
**New issue**. Write what you wish it did, in plain words. You do not need to
know how it would work. That is our job.

**Check back now and then.** New macros get added. Go to the
[Releases page](../../releases/latest), get the newest `.bas` file, and bring it
in again the same way you did in step 4. Your answers are kept.

**Give it a star if it helped.** Go to the
[project page](../../) and click the **Star** button near the top right. It is
free, and it helps other people find this.
