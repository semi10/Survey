package com.fsdragon.android.survey;

import android.app.AlertDialog;
import android.app.Dialog;
import android.app.DialogFragment;
import android.content.DialogInterface;
import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.widget.EditText;

/**
 * Created by Semion on 07/07/2015.
 */
public class NameIt extends DialogFragment {
    EditText name;
    @Override
    public Dialog onCreateDialog(Bundle savedInstanceState){


        AlertDialog.Builder builder = new AlertDialog.Builder(getActivity());
        LayoutInflater inflater = getActivity().getLayoutInflater();
        View view = inflater.inflate(R.layout.fragment_create, null);

        builder.setTitle(R.string.NameIt);
        builder.setView(view);
        name = (EditText) view.findViewById(R.id.name_et);
        builder.setNegativeButton(R.string.Cancel, new DialogInterface.OnClickListener() {
            @Override
            public void onClick(DialogInterface dialog, int which) {
                dismiss();
            }
        });

        builder.setPositiveButton(R.string.Finish, new DialogInterface.OnClickListener() {
            @Override
            public void onClick(DialogInterface dialog, int which) {
                String surveyName = name.getText().toString();
                ((createSurvey)getActivity()).create(surveyName);
                ((createSurvey)getActivity()).finish();
            }
        });
        Dialog dialog = builder.create();
        return dialog;
    }

}